--ERP库存备份
IF NOT EXISTS(SELECT 1 FROM SysObjects WHERE NAME = 't_Tous_StockImage' AND XTYPE = 'U')
CREATE TABLE t_Tous_StockImage
(
	FID				INT IDENTITY(1,1) NOT NULL,
	FDate			DATETIME NULL,
	FItemID			INT NOT NULL DEFAULT(0),
	FStockID		INT NOT NULL DEFAULT(0),
	FSPID			INT NOT NULL DEFAULT(0),
	FUPC			NVARCHAR(255) NOT NULL DEFAULT(''),
	FQty			DECIMAL(28,10) NOT NULL DEFAULT(0)
)
go

EXEC p_IC_AddIndex 'Idx_Tous_StockImage_FID','t_Tous_StockImage','FID'
EXEC p_IC_AddIndex 'Idx_Tous_StockImage_FDate','t_Tous_StockImage','FDate'
GO

--POS库存备份
IF NOT EXISTS(SELECT 1 FROM SysObjects WHERE NAME = 't_Tous_POSStockImage' AND XTYPE = 'U')
CREATE TABLE t_Tous_POSStockImage
(
	FID				INT IDENTITY(1,1) NOT NULL,
	FDate			DATETIME NULL,
	FItemID			INT NOT NULL DEFAULT(0),
	FStockID		INT NOT NULL DEFAULT(0),
	FSPID			INT NOT NULL DEFAULT(0),
	FUPC			NVARCHAR(255) NOT NULL DEFAULT(''),
	FQty			DECIMAL(28,10) NOT NULL DEFAULT(0)
)
go

EXEC p_IC_AddIndex 'Idx_Tous_StockImage_FID','t_Tous_POSStockImage','FID'
EXEC p_IC_AddIndex 'Idx_Tous_StockImage_FDate','t_Tous_POSStockImage','FDate'
GO
