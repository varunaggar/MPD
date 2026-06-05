<#
.SYNOPSIS
    Pester 5 test suite for the M365 Permissions Sync solution.

.DESCRIPTION
    Smoke tests, module validation, and unit tests for helper functions.
    These tests do NOT require a live tenant or SQL database — they
    validate code structure, module exports, and pure-logic functions.

    Run with:
      Invoke-Pester -Path .\M365PermSync.Tests.ps1 -Output Detailed

.NOTES
    Requires Pester v5+:  Install-Module Pester -MinimumVersion 5.0 -Force
#>

BeforeAll {
    $scriptRoot  = $PSScriptRoot
    $sharedPath  = Join-Path $scriptRoot "shared"
    $scriptsPath = $scriptRoot
}

# ══════════════════════════════════════════════════════════════
# Module Import Tests
# Validates that all shared modules can be imported without errors.
# This catches syntax errors, missing dependencies, and
# broken Export-ModuleMember declarations.
# ══════════════════════════════════════════════════════════════

Describe "Module Imports" {

    $modules = @(
        "ConfigHelpers",
        "LoggingHelpers",
        "SqlHelpers",
        "ExoHelpers",
        "GraphHelpers",
        "AdminApiHelpers",
        "DependencyHelpers"
    )

    foreach ($moduleName in $modules) {
        It "should import $moduleName without errors" {
            $modulePath = Join-Path $sharedPath "$moduleName.psm1"
            { Import-Module $modulePath -Force -ErrorAction Stop } | Should -Not -Throw
        }
    }
}

# ══════════════════════════════════════════════════════════════
# Exported Function Tests
# Validates that each module exports the expected public functions.
# ══════════════════════════════════════════════════════════════

Describe "SqlHelpers Exports" {
    BeforeAll {
        Import-Module (Join-Path $sharedPath "SqlHelpers.psm1") -Force
    }

    $expectedFunctions = @(
        "Initialize-SqlContext",
        "Open-SqlConnection",
        "Invoke-SqlNonQuery",
        "Invoke-SqlScalar",
        "Invoke-SqlQuery",
        "Get-DeltaToken",
        "Save-DeltaToken",
        "Disable-DeltaToken",
        "ConvertFrom-JwtToken"
    )

    foreach ($fn in $expectedFunctions) {
        It "should export $fn" {
            Get-Command $fn -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }
    }

    It "should NOT export private functions" {
        $exported = (Get-Module SqlHelpers).ExportedFunctions.Keys
        $exported | Should -Not -Contain "New-SqlConnection"
        $exported | Should -Not -Contain "Set-SqlParameters"
        $exported | Should -Not -Contain "Get-SqlAccessToken"
    }
}

Describe "GraphHelpers Exports" {
    BeforeAll {
        Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1") -Force
        Import-Module (Join-Path $sharedPath "GraphHelpers.psm1") -Force
    }

    $expectedFunctions = @(
        "Initialize-GraphContext",
        "Get-GraphToken",
        "Invoke-GraphRequest",
        "Invoke-GraphPagedRequest",
        "Invoke-GraphDeltaQuery",
        "Get-GraphMailFolderPaths",
        "Invoke-GraphHuntingQuery"
    )

    foreach ($fn in $expectedFunctions) {
        It "should export $fn" {
            Get-Command $fn -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }
    }
}

Describe "ExoHelpers Exports" {
    BeforeAll {
        Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1") -Force
        Import-Module (Join-Path $sharedPath "ExoHelpers.psm1") -Force
    }

    It "should export only connection management functions" {
        $exported = (Get-Module ExoHelpers).ExportedFunctions.Keys
        $exported | Should -Contain "Initialize-ExoContext"
        $exported | Should -Contain "Connect-ExoSession"
        $exported | Should -Contain "Disconnect-ExoSession"
        $exported.Count | Should -Be 3
    }

    It "should NOT export removed business-logic functions" {
        $exported = (Get-Module ExoHelpers).ExportedFunctions.Keys
        $exported | Should -Not -Contain "Get-AllExoMailboxes"
        $exported | Should -Not -Contain "Get-ChangedExoMailboxes"
        $exported | Should -Not -Contain "ConvertTo-MailboxObject"
    }
}

Describe "AdminApiHelpers Exports" {
    BeforeAll {
        Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1") -Force
        Import-Module (Join-Path $sharedPath "AdminApiHelpers.psm1") -Force
    }

    It "should export infrastructure functions" {
        $exported = (Get-Module AdminApiHelpers).ExportedFunctions.Keys
        $exported | Should -Contain "Initialize-AdminApiContext"
        $exported | Should -Contain "Get-AdminApiToken"
        $exported | Should -Contain "Get-AnchorMailboxHeader"
        $exported | Should -Contain "Invoke-AdminApiRequest"
        $exported | Should -Contain "Invoke-AdminApiPagedRequest"
    }

    It "should NOT export removed business-logic functions" {
        $exported = (Get-Module AdminApiHelpers).ExportedFunctions.Keys
        $exported | Should -Not -Contain "Get-MfcGroupMembers"
        $exported | Should -Not -Contain "Get-MailboxFolderPermissions"
    }
}

# ══════════════════════════════════════════════════════════════
# ConvertFrom-JwtToken Unit Tests
# ══════════════════════════════════════════════════════════════

Describe "ConvertFrom-JwtToken" {
    BeforeAll {
        Import-Module (Join-Path $sharedPath "SqlHelpers.psm1") -Force
    }

    It "should decode a valid JWT payload" {
        # JWT with payload: {"sub":"1234","name":"Test User","iat":1516239022}
        $jwt = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0IiwibmFtZSI6IlRlc3QgVXNlciIsImlhdCI6MTUxNjIzOTAyMn0.SflKxwRJSMeKKF2QT4fwpMeJf36POk6yJV_adQssw5c"
        $result = ConvertFrom-JwtToken -Token $jwt
        $result.sub  | Should -Be "1234"
        $result.name | Should -Be "Test User"
    }

    It "should throw on invalid JWT" {
        { ConvertFrom-JwtToken -Token "not.a.jwt.at.all" } | Should -Throw
    }

    It "should throw on empty string" {
        { ConvertFrom-JwtToken -Token "" } | Should -Throw
    }

    It "should handle base64 padding correctly" {
        # JWT with payload that needs padding: {"a":"b"}
        $jwt = "eyJhbGciOiJIUzI1NiJ9.eyJhIjoiYiJ9.signature"
        $result = ConvertFrom-JwtToken -Token $jwt
        $result.a | Should -Be "b"
    }
}

# ══════════════════════════════════════════════════════════════
# Script File Validation
# Validates that all scripts parse without syntax errors and
# have the expected structure (CmdletBinding, param block).
# ══════════════════════════════════════════════════════════════

Describe "Script Syntax Validation" {

    $scripts = @(
        "Invoke-UserBaselineLoad.ps1",
        "Invoke-UserDeltaSync.ps1",
        "Invoke-MailboxBaselineLoad.ps1",
        "Invoke-MailboxDeltaSync.ps1",
        "Invoke-PermissionBaselineLoad.ps1",
        "Invoke-MfcGroupDeltaSync.ps1",
        "Invoke-PermissionDeltaSync.ps1",
        "Invoke-TrusteeResolution.ps1"
    )

    foreach ($scriptName in $scripts) {
        Context $scriptName {

            BeforeAll {
                $scriptPath = Join-Path $scriptsPath $scriptName
                $content    = Get-Content $scriptPath -Raw -ErrorAction SilentlyContinue
            }

            It "should exist" {
                $scriptPath | Should -Exist
            }

            It "should parse without syntax errors" {
                $tokens = $null
                $errors = $null
                [System.Management.Automation.Language.Parser]::ParseFile(
                    $scriptPath, [ref]$tokens, [ref]$errors
                ) | Out-Null
                $errors.Count | Should -Be 0
            }

            It "should have CmdletBinding" {
                $content | Should -Match '\[CmdletBinding\(\)\]'
            }

            It "should have a ConfigPath parameter" {
                $content | Should -Match '\$ConfigPath'
            }

            It "should have comment-based help" {
                $content | Should -Match '\.SYNOPSIS'
            }
        }
    }
}

# ══════════════════════════════════════════════════════════════
# SQL Migration File Validation
# ══════════════════════════════════════════════════════════════

Describe "SQL Migration Files" {

    $sqlFiles = @(
        "01_shared_tables.sql",
        "02_users_table.sql",
        "03_mailboxes_table.sql",
        "04_permissions_tables.sql",
        "05_add_permissions_baselined.sql",
        "06_column_size_fixes.sql"
    )

    foreach ($sqlFile in $sqlFiles) {
        It "$sqlFile should exist" {
            Join-Path $scriptsPath $sqlFile | Should -Exist
        }
    }

    It "04_permissions_tables.sql should filter deleted mailboxes in view" {
        $content = Get-Content (Join-Path $scriptsPath "04_permissions_tables.sql") -Raw
        # Every Mailboxes join in the view should have IsDeleted = 0
        ($content | Select-String "m\.IsDeleted = 0" -AllMatches).Matches.Count | Should -BeGreaterOrEqual 4
        ($content | Select-String "mb\.IsDeleted = 0" -AllMatches).Matches.Count | Should -BeGreaterOrEqual 2
    }

    It "04_permissions_tables.sql should have filtered unique indexes" {
        $content = Get-Content (Join-Path $scriptsPath "04_permissions_tables.sql") -Raw
        ($content | Select-String "WHERE IsDeleted = 0" -AllMatches).Matches.Count | Should -BeGreaterOrEqual 4
    }

    It "06_column_size_fixes.sql should widen TokenValue to MAX" {
        $content = Get-Content (Join-Path $scriptsPath "06_column_size_fixes.sql") -Raw
        $content | Should -Match "TokenValue\s+NVARCHAR\(MAX\)"
    }

    It "06_column_size_fixes.sql should add NewMailboxesBaselined column" {
        $content = Get-Content (Join-Path $scriptsPath "06_column_size_fixes.sql") -Raw
        $content | Should -Match "NewMailboxesBaselined"
    }
}

# ══════════════════════════════════════════════════════════════
# Config.xml Structure Validation
# ══════════════════════════════════════════════════════════════

Describe "Config.xml Structure" {
    BeforeAll {
        $configPath = Join-Path $scriptsPath "config.xml"
        if (Test-Path $configPath) {
            [xml]$config = Get-Content $configPath
        }
    }

    It "should exist" {
        Join-Path $scriptsPath "config.xml" | Should -Exist
    }

    $requiredSections = @(
        "Authentication",
        "Database",
        "Logging",
        "Graph",
        "ExchangeOnline",
        "Sync"
    )

    foreach ($section in $requiredSections) {
        It "should have <$section> section" {
            $config.Config.$section | Should -Not -BeNullOrEmpty
        }
    }

    It "should have TenantId in Authentication" {
        $config.Config.Authentication.TenantId | Should -Not -BeNullOrEmpty
    }

    It "should have CertificateThumbprint in Authentication" {
        $config.Config.Authentication.CertificateThumbprint | Should -Not -BeNullOrEmpty
    }

    It "should have DeltaOverlapMinutes in Sync" {
        $config.Config.Sync.DeltaOverlapMinutes | Should -Not -BeNullOrEmpty
        [int]$config.Config.Sync.DeltaOverlapMinutes | Should -BeGreaterThan 0
    }

    It "should NOT have placeholder values" {
        $config.Config.Authentication.TenantId | Should -Not -BeLike "REPLACE*"
        $config.Config.Database.Server | Should -Not -BeLike "REPLACE*"
    }
}

# ══════════════════════════════════════════════════════════════
# Trustee Resolution Logic Unit Tests
# Tests the resolution matching tiers in isolation.
# ══════════════════════════════════════════════════════════════

Describe "Trustee Resolution Logic" {

    BeforeAll {
        # Simulate the lookup dictionaries
        $script:upnLookup = @{
            'john.smith@contoso.com' = 'guid-john'
            'jane.doe@contoso.com'   = 'guid-jane'
        }
        $script:mailLookup = @{
            'john.smith@contoso.com' = 'guid-john'
            'jsmith@contoso.com'     = 'guid-john'
        }
        $script:localPartLookup = @{
            'john.smith' = 'guid-john'
            'jane.doe'   = 'guid-jane'
            'duplicate'  = $null   # ambiguous
        }
        $script:displayNameLookup = @{
            'john smith' = 'guid-john'
            'common name'= $null   # ambiguous
        }

        # Inline resolution function for testing
        function Resolve-TestTrustee {
            param([string]$Identity)

            if ($Identity -in @('Default', 'Anonymous')) { return 'SYSTEM' }

            $lower = $Identity.ToLower().Trim()

            # Tier 1: UPN
            if ($script:upnLookup[$lower]) { return $script:upnLookup[$lower] }

            # Tier 2: Mail
            if ($script:mailLookup[$lower]) { return $script:mailLookup[$lower] }

            # Tier 3: DOMAIN\user
            if ($Identity -match '^[^\\@]+\\(.+)$') {
                $local = $Matches[1].ToLower()
                if ($script:localPartLookup.ContainsKey($local)) {
                    return $script:localPartLookup[$local]  # may be $null (ambiguous)
                }
            }

            # Tier 4: DN / path
            $displayCandidate = $null
            if ($Identity -match 'CN=([^,]+)') { $displayCandidate = $Matches[1].ToLower() }
            elseif ($Identity -match '/([^/]+)$') { $displayCandidate = $Matches[1].ToLower() }
            if ($displayCandidate -and $script:displayNameLookup.ContainsKey($displayCandidate)) {
                return $script:displayNameLookup[$displayCandidate]
            }

            return $null
        }
    }

    It "should resolve by exact UPN" {
        Resolve-TestTrustee 'john.smith@contoso.com' | Should -Be 'guid-john'
    }

    It "should resolve by exact Mail (different from UPN)" {
        Resolve-TestTrustee 'jsmith@contoso.com' | Should -Be 'guid-john'
    }

    It "should resolve DOMAIN\user format" {
        Resolve-TestTrustee 'CONTOSO\john.smith' | Should -Be 'guid-john'
    }

    It "should return null for ambiguous DOMAIN\user" {
        Resolve-TestTrustee 'CONTOSO\duplicate' | Should -BeNullOrEmpty
    }

    It "should resolve DN format (CN=...)" {
        Resolve-TestTrustee 'CN=John Smith,OU=Users,DC=contoso,DC=com' | Should -Be 'guid-john'
    }

    It "should resolve path format (org/ou/name)" {
        Resolve-TestTrustee 'contoso.com/Users/John Smith' | Should -Be 'guid-john'
    }

    It "should return null for ambiguous display name" {
        Resolve-TestTrustee 'CN=Common Name,OU=Users,DC=contoso,DC=com' | Should -BeNullOrEmpty
    }

    It "should return SYSTEM for Default trustee" {
        Resolve-TestTrustee 'Default' | Should -Be 'SYSTEM'
    }

    It "should return SYSTEM for Anonymous trustee" {
        Resolve-TestTrustee 'Anonymous' | Should -Be 'SYSTEM'
    }

    It "should return null for completely unknown identity" {
        Resolve-TestTrustee 'nobody@nowhere.com' | Should -BeNullOrEmpty
    }

    It "should be case-insensitive" {
        Resolve-TestTrustee 'JOHN.SMITH@CONTOSO.COM' | Should -Be 'guid-john'
    }
}

# ══════════════════════════════════════════════════════════════
# Cross-Reference Tests
# Validates consistency across scripts and schema.
# ══════════════════════════════════════════════════════════════

Describe "Cross-Reference Consistency" {

    It "all scripts should reference the same shared module set" {
        $scripts = Get-ChildItem $scriptsPath -Filter "Invoke-*.ps1"
        foreach ($s in $scripts) {
            $content = Get-Content $s.FullName -Raw
            # Every script must import ConfigHelpers and LoggingHelpers
            $content | Should -Match "ConfigHelpers"
            $content | Should -Match "LoggingHelpers"
        }
    }

    It "all scripts should use inline SyncLog INSERT (not Start-SyncLogEntry)" {
        $scripts = Get-ChildItem $scriptsPath -Filter "Invoke-*.ps1"
        foreach ($s in $scripts) {
            $content = Get-Content $s.FullName -Raw
            $content | Should -Not -Match "Start-SyncLogEntry"
            $content | Should -Not -Match "Complete-SyncLogEntry"
        }
    }

    It "no script should reference removed module functions" {
        $removed = @(
            "Invoke-LoggedPhase",
            "Invoke-SqlBatch",
            "Get-AllExoMailboxes",
            "Get-ChangedExoMailboxes",
            "ConvertTo-MailboxObject",
            "Get-MfcGroupMembers",
            "Get-MailboxFolderPermissions"
        )
        $scripts = Get-ChildItem $scriptsPath -Filter "Invoke-*.ps1"
        foreach ($s in $scripts) {
            $content = Get-Content $s.FullName -Raw
            foreach ($fn in $removed) {
                $content | Should -Not -Match $fn -Because "$($s.Name) should not reference removed function $fn"
            }
        }
    }

    It "PermissionBaselineLoad should seed permissions_delta_timestamp" {
        $content = Get-Content (Join-Path $scriptsPath "Invoke-PermissionBaselineLoad.ps1") -Raw
        $content | Should -Match "permissions_delta_timestamp"
    }

    It "PermissionDeltaSync Phase 6 should use Replace-FolderPermissions (re-fetch, not surgical)" {
        $content = Get-Content (Join-Path $scriptsPath "Invoke-PermissionDeltaSync.ps1") -Raw
        $content | Should -Match "Replace-FolderPermissions"
    }

    It "MailboxBaselineLoad and MailboxDeltaSync should check knownUserIds (FK fix)" {
        foreach ($scriptName in @("Invoke-MailboxBaselineLoad.ps1", "Invoke-MailboxDeltaSync.ps1")) {
            $content = Get-Content (Join-Path $scriptsPath $scriptName) -Raw
            $content | Should -Match "knownUserIds" -Because "$scriptName should have FK race condition fix"
        }
    }
}
