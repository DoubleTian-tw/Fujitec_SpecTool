CREATE TABLE [dbo].[BasicSetting_Table] (
    [ID]      INT   NOT NULL,
    [Designer]    TEXT NULL,
    [Checker] TEXT  NULL,
    [Approver] TEXT NULL, 
    [Local] TEXT NULL, 
    [MachineType] TEXT NULL, 
    [FLEX] TEXT NULL, 
    [OperationType] TEXT NULL, 
    PRIMARY KEY CLUSTERED ([ID] ASC)
);

