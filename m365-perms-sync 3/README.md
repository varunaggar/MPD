# M365 Permissions Sync

PowerShell-based solution running on Windows Server as scheduled tasks.
Maintains an Azure SQL database of M365 Exchange Online permissions,
starting with user information sync (Phase 1).

---

## Project structure

```
m365-perms-sync/
│
├── config.xml                          # All configuration — edit before first run
│
├── shared/                             # Reusable PowerShell modules (imported by every script)
│   ├── ConfigHelpers.psm1              # Reads config.xml, authenticates service principal
│   ├── LoggingHelpers.psm1             # File-based logging (ProcessName_yyyy-MM-dd_HH-mm-ss.log)
│   ├── GraphHelpers.psm1               # Graph REST API helpers (token, paging, delta queries)
│   └── SqlHelpers.psm1                 # Azure SQL helpers (AAD token auth, MERGE, SyncLog)
│
├── Invoke-UserBaselineLoad.ps1         # Full snapshot of all users — run once / weekly
├── Invoke-UserDeltaSync.ps1            # Incremental delta sync — runs every 15 min
│
├── sql/
│   ├── 01_shared_tables.sql            # DeltaTokens + SyncLog — deploy first
│   └── 02_users_table.sql              # Users table + indexes + views
│
├── logs/                               # Log files written here (auto-created)
│   └── (auto-generated)               # Invoke-UserBaselineLoad_2026-05-07_14-30-00.log
│
└── docs/
    └── data-model.html                 # Visual ER diagrams (open in browser)
```

---

## Architecture decisions

| Area | Decision |
|---|---|
| Compute | Windows Server scheduled tasks |
| Graph authentication | Certificate-based service principal (cert in Windows cert store) |
| SQL authentication | Azure AD token via same service principal (no SQL username/password) |
| Change detection | Graph /users/delta query — 15-min polling |
| Configuration | config.xml — all settings in one file |
| Logging | Individual log file per process: ProcessName_yyyy-MM-dd_HH-mm-ss.log |
| Deletion handling | Soft-delete (IsDeleted + DeletedAt) |
| Token safety | Delta token saved ONLY after all processing completes |

---

## Prerequisites

### 1. PowerShell 7 and Az module

```powershell
# Check PowerShell version (must be 7.x)
$PSVersionTable.PSVersion

# Install Az.Accounts module (required for Get-AzAccessToken)
Install-Module Az.Accounts -Scope AllUsers -Force
```

### 2. Certificate setup

```powershell
# Create self-signed certificate (or use your PKI)
$cert = New-SelfSignedCertificate `
    -Subject "CN=M365PermSync" `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Note the thumbprint
$cert.Thumbprint

# Export public key to upload to Entra app registration
Export-Certificate -Cert $cert -FilePath "C:\M365PermSync\M365PermSync.cer"
```

Upload `M365PermSync.cer` to your Entra app registration under
Certificates & secrets → Certificates → Upload certificate.

### 3. Entra app registration

Required Graph API permission (Application type — not Delegated):
- `User.Read.All`

Grant admin consent after adding the permission.

### 4. Azure SQL setup (run once as SQL admin)

```sql
-- Connect to db-m365permissions as SQL admin
-- Replace [YourAppDisplayName] with the display name of the Entra app registration
CREATE USER [YourAppDisplayName] FROM EXTERNAL PROVIDER;
ALTER ROLE db_datareader ADD MEMBER [YourAppDisplayName];
ALTER ROLE db_datawriter ADD MEMBER [YourAppDisplayName];
GRANT EXECUTE TO [YourAppDisplayName];
```

### 5. Service account permissions

The Windows service account running the scheduled tasks needs:
- Read access to the certificate private key in the Windows cert store
- Write access to the `logs\` directory
- Execute permission on PowerShell scripts

To grant cert private key access:
```
certlm.msc → Personal → Certificates → right-click cert
→ All Tasks → Manage Private Keys → add service account with Read
```

---

## Configuration (config.xml)

Edit `config.xml` before running anything. Replace all `REPLACE-WITH-*` values:

| Field | Where to find it |
|---|---|
| `TenantId` | Entra portal → Overview |
| `AppId` | Entra portal → App registrations → your app → Overview |
| `CertificateThumbprint` | `Get-ChildItem Cert:\LocalMachine\My` |
| `Database.Server` | Azure portal → SQL server → Server name |

---

## Deployment order

### Step 1 — Deploy SQL tables

Connect to `db-m365permissions` (not master) using SSMS or sqlcmd:

```bash
sqlcmd -S your-server.database.windows.net -d db-m365permissions ^
       --authentication-method ActiveDirectoryInteractive -G ^
       -i sql\01_shared_tables.sql

sqlcmd -S your-server.database.windows.net -d db-m365permissions ^
       --authentication-method ActiveDirectoryInteractive -G ^
       -i sql\02_users_table.sql
```

Verify: `SELECT name FROM sys.tables ORDER BY name;` should return
`DeltaTokens`, `SyncLog`, `Users`.

### Step 2 — Edit config.xml

Replace every `REPLACE-WITH-*` placeholder.

### Step 3 — Test authentication

```powershell
cd C:\M365PermSync
Import-Module .\shared\ConfigHelpers.psm1 -Force
$Config = Import-SyncConfig -Path .\config.xml
Connect-SyncServicePrincipal -Config $Config
# Should complete without error

Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/" | Select-Object -Expand Token | Measure-Object -Character
# Should return a character count > 500 (valid JWT token)
```

### Step 4 — Run the baseline load (once)

```powershell
pwsh -NonInteractive -File "C:\M365PermSync\Invoke-UserBaselineLoad.ps1"
```

Expected duration: 10–30 minutes for 40,000 users.
Check `logs\Invoke-UserBaselineLoad_*.log` for progress.

### Step 5 — Create scheduled tasks

**Baseline load — weekly (Sunday 02:00)**

```powershell
$action  = New-ScheduledTaskAction -Execute 'pwsh.exe' `
               -Argument '-NonInteractive -File "C:\M365PermSync\Invoke-UserBaselineLoad.ps1"'
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 02:00
$settings = New-ScheduledTaskSettingsSet `
               -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
               -MultipleInstances IgnoreNew `
               -StartWhenAvailable

Register-ScheduledTask `
    -TaskName "M365PermSync - User Baseline Load" `
    -Action   $action `
    -Trigger  $trigger `
    -Settings $settings `
    -RunLevel Highest
```

**Delta sync — every 15 minutes**

```powershell
$action  = New-ScheduledTaskAction -Execute 'pwsh.exe' `
               -Argument '-NonInteractive -File "C:\M365PermSync\Invoke-UserDeltaSync.ps1"'
$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 15) `
               -Once -At (Get-Date) -RepetitionDuration ([TimeSpan]::MaxValue)
$settings = New-ScheduledTaskSettingsSet `
               -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
               -MultipleInstances IgnoreNew `
               -StartWhenAvailable

Register-ScheduledTask `
    -TaskName "M365PermSync - User Delta Sync" `
    -Action   $action `
    -Trigger  $trigger `
    -Settings $settings `
    -RunLevel Highest
```

`MultipleInstances = IgnoreNew` prevents a second instance starting
if the previous run is still in progress.
`StartWhenAvailable` recovers a missed run on next trigger.

---

## Verifying the sync is working

```sql
-- Last 5 runs
SELECT TOP 5 FunctionName, StartedAt, CompletedAt, Status,
       UsersInserted, UsersUpdated, UsersSoftDeleted, ErrorCount
FROM SyncLog
ORDER BY StartedAt DESC;

-- Current user counts
SELECT
    COUNT(*)                                                    AS TotalRows,
    SUM(CASE WHEN IsDeleted = 0 THEN 1 ELSE 0 END)             AS ActiveUsers,
    SUM(CASE WHEN IsDeleted = 1 THEN 1 ELSE 0 END)             AS DeletedUsers,
    SUM(CASE WHEN UserType = 'Guest' AND IsDeleted = 0 THEN 1 ELSE 0 END) AS GuestUsers
FROM Users;

-- Delta token health
SELECT TokenName, IsActive, UpdatedAt, DeactivationReason
FROM DeltaTokens;

-- Freshness check — last successful delta run
SELECT TOP 1 StartedAt, CompletedAt, Status, UsersProcessed
FROM SyncLog
WHERE FunctionName = 'Invoke-UserDeltaSync'
  AND Status = 'Success'
ORDER BY StartedAt DESC;
```

---

## Log files

Every script creates its own log file in the `logs\` directory:

```
logs\
├── Invoke-UserBaselineLoad_2026-05-07_02-00-00.log
├── Invoke-UserDeltaSync_2026-05-07_14-00-00.log
├── Invoke-UserDeltaSync_2026-05-07_14-15-00.log
└── ...
```

Log files older than `Logging.RetentionDays` (default 30) are deleted
automatically at the end of each script run.

---

## What's next (Phase 2)

After user sync is validated:
1. `Invoke-MailboxBaselineLoad.ps1` — all mailboxes via EXO PowerShell
2. `Invoke-MailboxDeltaSync.ps1` — incremental via WhenChangedUTC filter
3. `sql/03_mailboxes_table.sql` — Mailboxes table DDL
4. Extend SyncLog counter columns for mailbox metrics
