------------------------------------------------------------
--Craete Date:2018.09.23
--Last Modify:
--Description:给核算项目增加字段
------------------------------------------------------------

--*****************************  物料  **************************************
--为核算项目物料新增字段核算项目详细（t_ICItemCustom）--

--EAN 8
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FEAN8')
	ALTER TABLE t_ICItemCustom ADD FEAN8 NVARCHAR(255) 
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FEAN8'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'EAN 8','EAN 8',N'EAN 8','FEAN8',200,255,0,200,255,1,0,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,'' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FEAN8'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FEAN8','STRING',N'EAN 8','EAN 8',N'EAN 8',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO

--EAN 13
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FEAN13')
	ALTER TABLE t_ICItemCustom  ADD FEAN13 NVARCHAR(255)
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FEAN13'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'EAN 13','EAN 13',N'EAN 13','FEAN13',200,255,0,200,255,1,0,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,'' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FEAN13'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FEAN13','STRING',N'EAN 13','EAN 13',N'EAN 13',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO

--HQ Currency
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FHQCurrency')
	ALTER TABLE t_ICItemCustom ADD FHQCurrency NVARCHAR(255)
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FHQCurrency'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'HQ Currency','HQ Currency',N'HQ Currency','FHQCurrency',200,255,0,200,255,1,0,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,'' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FHQCurrency'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FHQCurrency','STRING',N'HQ Currency','HQ Currency',N'HQ Currency',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO

--HQ Cost Price
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FHQCostPrice')
	ALTER TABLE t_ICItemCustom ADD FHQCostPrice FLOAT
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FHQCostPrice'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'HQ Cost Price','HQ Cost Price',N'HQ Cost Price','FHQCostPrice',5,53,0,5,53,1,0,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,'' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FHQCostPrice'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FHQCostPrice','STRING',N'HQ Cost Price','HQ Cost Price',N'HQ Cost Price',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO

--HQ Retail Price
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FHQRetailPrice')
	ALTER TABLE t_ICItemCustom ADD FHQRetailPrice FLOAT
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FHQRetailPrice'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'HQ Sales Price','HQ Sales Price',N'HQ Sales Price','FHQRetailPrice',5,53,0,5,53,1,0,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,'' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FHQRetailPrice'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FHQRetailPrice','STRING',N'HQ Sales Price','HQ Sales Price',N'HQ Sales Price',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO

--Updated Flag
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FUpdatedFlag')
BEGIN
	ALTER TABLE t_ICItemCustom ADD FUpdatedFlag INT not null Default((1))
END

	
DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FUpdatedFlag'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'Updated Flag','Updated Flag',N'Updated Flag','FUpdatedFlag',3,0,0,3,0,null,0,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition','',0,16,'0' FROM t_ItemPropDesc WHERE FItemClassID = 4


--MacroFamily_OLD
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FMacroFamily')
	ALTER TABLE t_ICItemCustom ADD FMacroFamily NVARCHAR(255) 
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FMacroFamily'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'MacroFamily_OLD','MacroFamily_OLD',N'MacroFamily_OLD','FMacroFamily',200,255,0,200,255,1,0,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,'' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FMacroFamily'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FMacroFamily','STRING',N'MacroFamily_OLD','MacroFamily_OLD',N'MacroFamily_OLD',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO

--Family_OLD
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FFamily')
	ALTER TABLE t_ICItemCustom ADD FFamily NVARCHAR(255) 
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FFamily'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'Family_OLD','Family_OLD',N'Family_OLD','FFamily',200,255,0,200,255,1,0,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,'' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FFamily'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FFamily','STRING',N'Family_OLD','Family_OLD',N'Family_OLD',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO

--Type_OLD
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FTSType')
	ALTER TABLE t_ICItemCustom ADD FTSType NVARCHAR(255) 
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FTSType'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'Type_OLD','Type_OLD',N'Type_OLD','FTSType',200,255,0,200,255,1,0,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,'' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FTSType'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FTSType','STRING',N'Type_OLD','Type_OLD',N'Type_OLD',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO


--Metal_OLD
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FMetal')
	ALTER TABLE t_ICItemCustom ADD FMetal NVARCHAR(255) 
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FMetal'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'Metal_OLD','Metal_OLD',N'Metal_OLD','FMetal',200,255,0,200,255,1,0,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,'' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FMetal'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FMetal','STRING',N'Metal_OLD','Metal_OLD',N'Metal_OLD',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO


--macrofamily
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FMacroFamilyID')
	ALTER TABLE t_ICItemCustom ADD FMacroFamilyID int
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FMacroFamilyID'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue,FSrcTable,FSrcField,FDisplayField) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'MacroFamily','MacroFamily',N'MacroFamily','FMacroFamilyID',3,0,0,3,0,'',3001,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,0,'t_Item_3001','FItemID','FName' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FMacroFamilyID'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FMacroFamilyID','INTEGER',N'MacroFamily','MacroFamily',N'MacroFamily',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO

--family
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FFamilyID')
	ALTER TABLE t_ICItemCustom ADD FFamilyID int
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FFamilyID'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue,FSrcTable,FSrcField,FDisplayField) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'Family','Family',N'Family','FFamilyID',3,0,0,3,0,'',3002,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,0,'t_Item_3002','FItemID','FName' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FFamilyID'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FFamilyID','INTEGER',N'Family','Family',N'Family',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO

--TSType
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FTSTypeID')
	ALTER TABLE t_ICItemCustom ADD FTSTypeID int
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FTSTypeID'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue,FSrcTable,FSrcField,FDisplayField) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'Type','Type',N'Type','FTSTypeID',3,0,0,3,0,'',3003,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,0,'t_Item_3003','FItemID','FName' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FTSTypeID'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FTSTypeID','INTEGER',N'Type','Type',N'Type',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO


--Metal
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FMetalID')
	ALTER TABLE t_ICItemCustom ADD FMetalID int
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FMetalID'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue,FSrcTable,FSrcField,FDisplayField) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'Metal','Metal',N'Metal','FMetalID',3,0,0,3,0,'',3004,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,0,'t_Item_3004','FItemID','FName' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FMetalID'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FMetalID','INTEGER',N'Metal','Metal',N'Metal',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO

--Collection
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_ICItemCustom' AND t2.Name = 'FCollectionID')
	ALTER TABLE t_ICItemCustom ADD FCollectionID int
GO

DELETE FROM t_ItemPropDesc WHERE FItemClassID = 4 AND FSQLColumnName = 'FCollectionID'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue,FSrcTable,FSrcField,FDisplayField) 
SELECT 4,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'Collection','Collection',N'Collection','FCollectionID',3,0,0,3,0,'',3005,N'90.Tous Definition',N'90.Tous Definition',N'90.Tous Definition',NULL,0,0,0,'t_Item_3005','FItemID','FName' FROM t_ItemPropDesc WHERE FItemClassID = 4

DELETE FROM t_FieldDescription WHERE FTableID = (SELECT FTableID FROM t_TableDescription WHERE FTableName = 't_ICItem') AND FFieldName = 'FCollectionID'
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FDescription_CHT,FDescription_EN,FSefDefSign) 
SELECT FTableID,'FCollectionID','INTEGER',N'Collection','Collection',N'Collection',0 FROM t_TableDescription WHERE FTableName = 't_ICItem'
GO


EXEC SP_CREATE_ICItem_VIEW 
EXEC SP_CREATE_ICItem_TRIGGER
GO


--*****************************  仓库  **************************************
---------------------------为为核算项目仓库新增1个字段，放在“Custom”页签:核算项目详细（t_Stock）-----------------------------------
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_Stock' AND t2.Name = 'FInventoryType')
	ALTER TABLE t_Stock ADD FInventoryType INT 
GO
DELETE FROM t_ItemPropDesc WHERE FItemClassID = 5 AND FSQLColumnName = 'FInventoryType'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FSrcTable,FSrcField,FDisplayField,FDefaultValue,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FOrder) 
VALUES (5,101,N'库存类型',N'庫存類型',N'Inventory Type_OLD','FInventoryType',3,0,0,3,0,NULL,11200,'t_SubMessage','FInterID','FName','',N'01.基本资料',N'01.基本資料',N'01.Basic Data',NULL,1,2,0)
GO

--*****************************  客户  **************************************
---------------------------为核算项目客户新增"自动生成发票"，放在“基本资料”页签:核算项目详细（t_Organization）-----------------------------------
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_Organization' AND t2.Name = 'FAutoInv')
	ALTER TABLE t_Organization ADD FAutoInv bit 
GO
DELETE FROM t_ItemPropDesc WHERE FItemClassID = 1 AND FSQLColumnName = 'FAutoInv'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 1,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'自动生成发票','自動生成發票',N'Auto Generate Invoice','FAutoInv',11,0,0,11,0,1,0,N'01.基本资料',N'01.基本資料',N'01.Basic Data','',0,0,'1' FROM t_ItemPropDesc WHERE FItemClassID = 1

---------------------------为核算项目客户新增"是否经销商门店"，放在“基本资料”页签:核算项目详细（t_Organization）-----------------------------------
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_Organization' AND t2.Name = 'FIsFranchiser')
	ALTER TABLE t_Organization ADD FIsFranchiser bit 
GO
DELETE FROM t_ItemPropDesc WHERE FItemClassID = 1 AND FSQLColumnName = 'FIsFranchiser'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 1,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'是否经销商门店','是否经销商门店',N'Franchiser Store','FIsFranchiser',11,0,0,11,0,1,0,N'01.基本资料',N'01.基本資料',N'01.Basic Data','',0,0,'1' FROM t_ItemPropDesc WHERE FItemClassID = 1

---------------------------经销商，放在“基本资料”页签:核算项目详细（t_Organization）-----------------------------------
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_Organization' AND t2.Name = 'FFranchiserID')
	ALTER TABLE t_Organization ADD FFranchiserID INT 
GO
DELETE FROM t_ItemPropDesc WHERE FItemClassID = 1 AND FSQLColumnName = 'FFranchiserID'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FSrcTable,FSrcField,FDisplayField,FDefaultValue,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FOrder) 
VALUES (1,101,N'经销商',N'經銷商',N'Franchiser','FFranchiserID',3,0,0,3,0,NULL,11500,'t_SubMessage','FInterID','FName','',N'01.基本资料',N'01.基本資料',N'01.Basic Data',NULL,1,2,0)
GO
---------------------------为核算项目客户新增"是否已导出给POS"，放在“基本资料”页签:核算项目详细（t_Organization）-----------------------------------
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 't_Organization' AND t2.Name = 'FIsExported')
	ALTER TABLE t_Organization ADD FIsExported INT
GO
DELETE FROM t_ItemPropDesc WHERE FItemClassID = 1 AND FSQLColumnName = 'FIsExported'
INSERT INTO t_ItemPropDesc (FItemClassID,FPropID,FOrder,FName,FName_en,FName_cht,FSQLColumnName,FDataType,FPrecision,FScale,FActualType,FActualSize,FBehavior,FSearch,FPageName,FPageName_en,FPageName_cht,FComCall,FIsShownList,FViewMask,FDefaultValue) 
SELECT 1,ISNULL(MAX(FPropID),0)+1,ISNULL(MAX(FOrder),0)+1,N'导出标识','導出標識',N'Export Flag','FIsExported',11,0,0,11,0,1,0,N'01.基本资料',N'01.基本資料',N'01.Basic Data','',0,16,'1' FROM t_ItemPropDesc WHERE FItemClassID = 1


--*****************************  更新视图  **************************************
EXEC SP_CREATE_ICItem_VIEW 
EXEC SP_CREATE_ICItem_TRIGGER
GO







