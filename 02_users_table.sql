-- =============================================================
-- 02_users_table.sql
-- Users table — destination for Graph /users baseline + delta sync.
-- Deploy ONCE after 01_shared_tables.sql.
-- =============================================================

IF OBJECT_ID('Users', 'U') IS NULL
BEGIN
    CREATE TABLE Users (
        -- Immutable primary key from Entra ID (GUID)
        UserId                UNIQUEIDENTIFIER NOT NULL PRIMARY KEY,

        -- Core identity attributes (can change over time)
        UserPrincipalName     NVARCHAR(256)    NOT NULL,
        DisplayName           NVARCHAR(256)    NULL,
        Mail                  NVARCHAR(256)    NULL,

        -- Account state
        AccountEnabled        BIT              NOT NULL DEFAULT 1,
        UserType              NVARCHAR(20)     NULL,   -- Member | Guest
        OnPremisesSyncEnabled BIT              NULL,   -- NULL = cloud-only

        -- Organisational attributes (for reporting / RLS later)
        Department            NVARCHAR(128)    NULL,
        JobTitle              NVARCHAR(128)    NULL,

        -- Lifecycle timestamps
        EntraCreatedDateTime  DATETIME2        NULL,   -- when Entra created the user
        FirstSeenAt           DATETIME2        NOT NULL DEFAULT SYSUTCDATETIME(),
        LastSyncedAt          DATETIME2        NOT NULL DEFAULT SYSUTCDATETIME(),
        LastModifiedAt        DATETIME2        NULL,   -- when we last detected a change

        -- Soft-delete flags
        IsDeleted             BIT              NOT NULL DEFAULT 0,
        DeletedAt             DATETIME2        NULL,

        -- Audit trail
        SyncSource            NVARCHAR(20)     NOT NULL,   -- Baseline | Delta
        LastSyncRunId         UNIQUEIDENTIFIER NULL
    );

    -- Filtered indexes — most lookups care only about active users
    CREATE INDEX IX_Users_UPN          ON Users(UserPrincipalName) WHERE IsDeleted = 0;
    CREATE INDEX IX_Users_Mail         ON Users(Mail)              WHERE IsDeleted = 0;
    CREATE INDEX IX_Users_IsDeleted    ON Users(IsDeleted, DeletedAt);
    CREATE INDEX IX_Users_LastSyncedAt ON Users(LastSyncedAt);

    PRINT 'Created Users table and indexes';
END;
GO
