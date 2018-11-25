--Import STN Temp Data Table
IF NOT EXISTS(SELECT 1 FROM SysObjects WHERE NAME = 't_Tous_STNImportData' AND XTYPE = 'U')
CREATE TABLE t_Tous_STNImportData
(
	FID				INT IDENTITY(1,1) NOT NULL,
	FStockOutID		INT NOT NULL,
	FStockInID		INT NOT NULL,
	FItemID			INT NOT NULL,
	FQty			DECIMAL(28,10) NOT NULL DEFAULT(0),
	FUUID			nvarchar(255) NOT NULL
)
go

EXEC p_IC_AddIndex 'Idx_Tous_STNImportData_FID','t_Tous_STNImportData','FID'
GO
