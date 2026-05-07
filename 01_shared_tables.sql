-- =============================================================
-- 01_shared_tables.sql
-- Shared operational tables used by every function in the solution.
-- Deploy ONCE during initial database setup.
-- =============================================================

IF OBJECT_ID('DeltaTokens', 'U') IS NULL
BEGIN
    CREATE TABLE DeltaTokens (
        TokenName          NVARCHAR(100)  NOT NULL PRIMARY KEY,
        TokenValue         NVARCHAR(2000) NOT NULL,
        CreatedAt          DATETIME2      NOT NULL DEFAULT SYSUTCDATETIME(),
        UpdatedAt          DATETIME2      NOT NULL DEFAULT SYSUTCDATETIME(),
        IsActive           BIT            NOT NULL DEFAULT 1,
        DeactivatedAt      DATETIME2      NULL,
        DeactivationReason NVARCHAR(500)  NULL
    );
    PRINT 'Created DeltaTokens table';
END;
GO

IF OBJECT_ID('SyncLog', 'U') IS NULL
BEGIN
    CREATE TABLE SyncLog (
        RunId            UNIQUEIDENTIFIER NOT NULL PRIMARY KEY DEFAULT NEWID(),
        FunctionName     NVARCHAR(100)    NOT NULL,
        StartedAt        DATETIME2        NOT NULL,
        CompletedAt      DATETIME2        NULL,
        Status           NVARCHAR(20)     NULL,    -- Running | Success | PartialFailure | Failed
        UsersInserted    INT              NOT NULL DEFAULT 0,
        UsersUpdated     INT              NOT NULL DEFAULT 0,
        UsersSoftDeleted INT              NOT NULL DEFAULT 0,
        UsersProcessed   INT              NOT NULL DEFAULT 0,
        ErrorCount       INT              NOT NULL DEFAULT 0,
        ErrorMessage     NVARCHAR(2000)   NULL,
        TokenAdvancedTo  NVARCHAR(2000)   NULL
    );

    CREATE INDEX IX_SyncLog_FunctionName_StartedAt
        ON SyncLog(FunctionName, StartedAt DESC);
    CREATE INDEX IX_SyncLog_Status_StartedAt
        ON SyncLog(Status, StartedAt DESC);
    PRINT 'Created SyncLog table and indexes';
END;
GO
