---***************  盘点方案增加字段 *****************************
IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 'icstockcheckprocess' AND t2.Name = 'FIsExported')
	ALTER TABLE icstockcheckprocess ADD FIsExported int
GO

IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 'icstockcheckprocess' AND t2.Name = 'FIsExportedToHQ')
	ALTER TABLE icstockcheckprocess ADD FIsExportedToHQ int
GO

---***************  销售出库单增加字段 *****************************
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=76
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='B01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='B01'

--***********  表头增加字段“POS Number”'FPOSNumber'  ********************
DELETE FROM ICTemplate WHERE FID='B01' AND FFieldName='FPOSNumber'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('B01',@CtlIndex,61,'POS单号:','POS單號:','POS Number:',0,-1,1,0,0,'FPOSNumber',
2895,720,2580,480,0,0,0,
'宋体',9,0,0,'','FPOSNumber','',0,3,
0,31,'kdtext','',0,
'','','',0,0)

--万能表增加字段 FPOSNumber
DELETE FROM t_FieldDescription WHERE FTableID=200004 AND FFieldName='FPOSNumber' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(200004,'FPOSNumber','STRING','POS单号','POS单号',0,0,'POS單號','POS Number','POS單號','POS Number')

--序事薄增加字段‘FPOSNumber’
DELETE FROM ICChatBillTitle WHERE FTypeID=76 AND FColName='FPOSNumber'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,76,'POS单号$','','FPOSNumber','ICStockBill',1,1000,1,-1,1,0,1,0,'FPOSNumber','v1','',0,0,
	'',1,0,1,'POS單號$','POS Number$',0,0,5,'',0,0,'',0)

--套打关系增加‘FPOSNumber’
DELETE FROM GLNotecitation WHERE FTemplateID='B01' AND FCode='FPOSNumber'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('B01',@MaxFID,50,'FPOSNumber','POS Number$',0,0,@CtlIndex,0,'FPOSNumber','','',1,'FPOSNumber$','POS Number$')
END
GO

EXEC p_AddStockColumn 'FPOSNumber',1,0
go

-----*************************************  销售发票 Sales Invoice  *************************************
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=62
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='I04'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='I04'
--表头增加字段“POS Number”'FPOSNumber'
DELETE FROM ICTemplate WHERE FID='I04' AND FFieldName='FPOSNumber'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('I04',@CtlIndex,62,'POS单号:','POS單號:','POS Number:',0,-1,1,0,0,'FPOSNumber',
2895,720,2580,480,0,0,0,
'宋体',9,0,0,'','FPOSNumber','',0,3,
0,31,'kdtext','',0,
'','','',0,0)


--万能表增加字段 FPOSNumber
DELETE FROM t_FieldDescription WHERE FTableID=200004 AND FFieldName='FPOSNumber' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(200004,'FPOSNumber','STRING','POS单号','POS单号',0,0,'POS單號','POS Number','POS單號','POS Number')

--序事薄增加字段‘FPOSNumber’
DELETE FROM ICChatBillTitle WHERE FTypeID=62 AND FColName='FPOSNumber'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,62,'POS单号$','','FPOSNumber','ICSale',1,1000,1,-1,1,0,1,0,'FPOSNumber','v1','',0,0,
	'',1,0,1,'POS單號$','POS Number$',0,0,5,'',0,0,'',0)


--套打关系增加‘FPOSNumber’
DELETE FROM GLNotecitation WHERE FTemplateID='I04' AND FCode='FPOSNumber'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('I04',@MaxFID,50,'FPOSNumber','POS单号$',0,0,@CtlIndex,0,'FPOSNumber','','',1,'Pos Number$','POS單號$')
END
GO

IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 'ICSale' AND t2.Name = 'FPOSNumber')
	ALTER TABLE ICSale ADD FPOSNumber varchar(255)  default ''
GO


--******************************  采购订单  **************************************************
--采购订单增加字段表头FPurchaseType "Purchase Type"
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=81
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='P02'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='P02'

--表头增加字段“Purchase Type”'FPurchaseType'
DELETE FROM ICTemplate WHERE FID='P02' AND FFieldName='FPurchaseType'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('P02',@CtlIndex,81,'采购类型:','採購類型:','Purchase Type:',2,11100,1,1,1,'FPurchaseType',
2895,720,2580,480,48,0,0,
'宋体',9,0,0,'','','',0,3,
0,31,'kdtext','',0,
0,'','',2,0)


--万能表增加字段 FPurchaseType
DELETE FROM t_FieldDescription WHERE FTableID=200004 AND FFieldName='FPurchaseType' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(200004,'FPurchaseType','STRING','Purchase Type','Purchase Type',0,0,'Purchase Type','FPurchaseType','Purchase Type','Purchase Type')

--序事薄增加字段‘FPurchaseType’
--DELETE FROM ICChatBillTitle WHERE FTypeID=81 AND FColName='FPurchaseType'
--INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
--	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
--	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
--	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
--	FFormat,FFormatType)
--VALUES(@MaxInterID,81,'采购类型$','','FPurchaseType','t_SubMessage',1,1000,1,-1,1,0,1,0,'FName','v1','',0,0,
--	'',1,0,1,'採購類型$','FPurchaseType$',0,0,5,'',0,0,'',0)


--套打关系增加‘FPurchaseType’
DELETE FROM GLNotecitation WHERE FTemplateID='P02' AND FCode='FPurchaseType'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('P02',@MaxFID,50,'FPurchaseType','Purchase Type$',0,0,@CtlIndex,0,'FPurchaseType','','',1,'FPurchaseType$','Purchase Type$')
END
GO

IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 'POOrder' AND t2.Name = 'FPurchaseType')
	ALTER TABLE POOrder ADD FPurchaseType int
GO


--采购订单增加字段表头FIsExport
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=81
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='P02'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='P02'

--表头增加字段“Is Export”'FIsExport'
DELETE FROM ICTemplate WHERE FID='P02' AND FFieldName='FIsExported'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,FMaxValue,FMinValue,FLookUpType,FSysMustInputItem,FAllowCopy)
Values('P02',@CtlIndex,81,'导出标识','導出標識','Export Flag',33,-1,1,3,0,'FIsExported',
4185,885,2580,480,48,0,0,
'宋体',9,0,0,'','','',0,3,
0,0,'kdtext','',0,
'0','','',0,0,1)

END
GO

IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 'POOrder' AND t2.Name = 'FIsExported')
	ALTER TABLE POOrder ADD FIsExported int
GO

--采购订单增加字段表头FHQNumber
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=81
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='P02'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='P02'
--表头增加字段“HQ Number”'FHQNumber'
DELETE FROM ICTemplate WHERE FID='P02' AND FFieldName='FHQNumber'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('P02',@CtlIndex,81,'HQ单号:','HQ單號:','HQ Number:',0,-1,1,0,0,'FHQNumber',
2895,720,2580,480,0,0,0,
'宋体',9,0,0,'','FHQNumber','',0,3,
0,31,'kdtext','',0,
'','','',0,0)


--万能表增加字段 FHQNumber
DELETE FROM t_FieldDescription WHERE FTableID=200004 AND FFieldName='FHQNumber' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(200004,'FHQNumber','STRING','HQ单号','HQ单号',0,0,'HQ單號','HQ Number','HQ單號','HQ Number')

--序事薄增加字段‘FHQNumber’
DELETE FROM ICChatBillTitle WHERE FTypeID=81 AND FColName='FHQNumber'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,81,'HQ单号$','','FHQNumber','POOrder',1,1000,1,-1,1,0,1,0,'FHQNumber','v1','',0,0,
	'',1,0,1,'HQ單號$','HQ Number$',0,0,5,'',0,0,'',0)


--套打关系增加‘FHQNumber’
DELETE FROM GLNotecitation WHERE FTemplateID='P02' AND FCode='FHQNumber'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('P02',@MaxFID,50,'FHQNumber','HQ单号$',0,0,@CtlIndex,0,'FHQNumber','','',1,'HQ Number$','HQ單號$')
END
GO

IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 'POOrder' AND t2.Name = 'FHQNumber')
	ALTER TABLE POOrder ADD FHQNumber varchar(255)  default ''
GO

--采购订单增加字段表头FHQOrigin
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=81
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='P02'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='P02'
--表头增加字段“Origin Code”'FHQOrigin'
DELETE FROM ICTemplate WHERE FID='P02' AND FFieldName='FHQOrigin'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('P02',@CtlIndex,81,'Origin Code:','Origin Code:','Origin Code:',0,-1,1,0,0,'FHQOrigin',
2895,720,2580,480,0,0,0,
'宋体',9,0,0,'','FHQOrigin','',0,3,
0,31,'kdtext','',0,
'','','',0,0)


--万能表增加字段 FHQOrigin
DELETE FROM t_FieldDescription WHERE FTableID=200004 AND FFieldName='FHQOrigin' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(200004,'FHQOrigin','STRING','Origin Code','Origin Code',0,0,'Origin Code','Origin Code','Origin Code','Origin Code')

--序事薄增加字段‘FHQOrigin’
DELETE FROM ICChatBillTitle WHERE FTypeID=81 AND FColName='FHQOrigin'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,81,'Origin Code$','','FHQOrigin','POOrder',1,1000,1,-1,1,0,1,0,'FHQOrigin','v1','',0,0,
	'',1,0,1,'Origin Code$','Origin Code$',0,0,5,'',0,0,'',0)


--套打关系增加‘FHQOrigin’
DELETE FROM GLNotecitation WHERE FTemplateID='P02' AND FCode='FHQOrigin'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('P02',@MaxFID,50,'FHQOrigin','Origin Code$',0,0,@CtlIndex,0,'FHQOrigin','','',1,'Origin Code$','Origin Code$')
END
GO

IF NOT EXISTS(SELECT 1 FROM SysObjects t1 ,SysColumns t2 WHERE t1.id = t2.id AND t1.name = 'POOrder' AND t2.Name = 'FHQOrigin')
	ALTER TABLE POOrder ADD FHQOrigin varchar(255)  default ''
GO
--******************************  外购入库单  **************************************************
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=58
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='A01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='A01'

--表头增加字段“Purchase Type”'FPOSNumber'
DELETE FROM ICTemplate WHERE FID='A01' AND FFieldName='FPOSNumber'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,FLookUpCls,FNeedSave,FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('A01',@CtlIndex,58,'POS单号:','POS單號:','POS Number:',
0,-1,1,0,0,'FPOSNumber',2855,
1135,2250,330,48,0,0,
'宋体',9,0,0,'','FPOSNumber','',0,3,
0,31,'kdtext','',0,
'','','',0,0)

--SELECT * FROM ICClassTableInfo  where FClassTypeID=1000002 and FCaption_CHS like '%单位%'
--SELECT * FROM ICClassType WHERE FName_CHS LIKE '%发票%'

--万能表增加字段 FPOSNumber
DELETE FROM t_FieldDescription WHERE FTableID=210008 AND FFieldName='FPOSNumber' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(210008,'FPOSNumber','STRING','POS单号','POS单号',0,0,'POS單號','POS Number','POS單號','POS Number')

--序事薄增加字段‘FPOSNumber’
DELETE FROM ICChatBillTitle WHERE FTypeID=58 AND FColName='FPOSNumber'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,58,'POS单号$','','FPOSNumber','ICStockBill',1,1000,1,-1,1,0,1,0,'FPOSNumber','v1','',0,0,
	'',1,0,1,'POS單號$','POS Number$',0,0,5,'',0,0,'',0)


--套打关系增加‘FPOSNumber’
DELETE FROM GLNotecitation WHERE FTemplateID='A01' AND FCode='FPOSNumber'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('A01',@MaxFID,19,'FPOSNumber','POS单号$',0,0,@CtlIndex,0,'FPOSNumber','','',1,'POS Number$','POS單號$')
END
GO

--外购入库单增加字段表头FIsExportedToHQ
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

--SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=81
--SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='C01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='A01'

--表头增加字段“Is Export”'FIsExport'
DELETE FROM ICTemplate WHERE FID='A01' AND FFieldName='FIsExportedToHQ'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,FMaxValue,FMinValue,FLookUpType,FSysMustInputItem,FAllowCopy)
Values('A01',@CtlIndex,58,'HQ导出标识','HQ導出標識','HQ Export Flag',33,-1,1,3,0,'FIsExportedToHQ',
4185,885,2580,480,48,0,0,
'宋体',9,0,0,'','','',0,3,
0,0,'kdtext','',0,
'0','','',0,0,1)

END
GO

--外购入库单增加字段表头FHQNumber
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=58
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='A01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='A01'
--表头增加字段“HQ Number”'FHQNumber'
DELETE FROM ICTemplate WHERE FID='A01' AND FFieldName='FHQNumber'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('A01',@CtlIndex,58,'HQ单号:','HQ單號:','HQ Number:',0,-1,1,0,0,'FHQNumber',
2895,720,2580,480,0,0,0,
'宋体',9,0,0,'','FHQNumber','',0,3,
0,31,'kdtext','',0,
'','','',0,0)


--万能表增加字段 FHQNumber
DELETE FROM t_FieldDescription WHERE FTableID=210008 AND FFieldName='FHQNumber' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(210008,'FHQNumber','STRING','HQ单号','HQ单号',0,0,'HQ單號','HQ Number','HQ單號','HQ Number')

--序事薄增加字段‘FHQNumber’
DELETE FROM ICChatBillTitle WHERE FTypeID=81 AND FColName='FHQNumber'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,58,'HQ单号$','','FHQNumber','ICStockBill',1,1000,1,-1,1,0,1,0,'FHQNumber','v1','',0,0,
	'',1,0,1,'HQ單號$','HQ Number$',0,0,5,'',0,0,'',0)


--套打关系增加‘FHQNumber’
DELETE FROM GLNotecitation WHERE FTemplateID='A01' AND FCode='FHQNumber'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('A01',@MaxFID,50,'FHQNumber','HQ单号$',0,0,@CtlIndex,0,'FHQNumber','','',1,'HQ Number$','HQ單號$')
END
GO

EXEC p_AddStockColumn 'FHQNumber',1,0
go

--******************************  其他出库单  **************************************************
--其他出库单增加字段表头FAdjReason "Adjustment Reason"
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=68
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='B09'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='B09'

--表头增加字段“Adjustment Reason”'FAdjReason'
DELETE FROM ICTemplate WHERE FID='B09' AND FFieldName='FAdjReason'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('B09',@CtlIndex,61,'调整原因:','調整原因:','Adjustment Reason:',2,11400,1,1,1,'FAdjReason',
2895,720,2580,480,48,0,0,
'宋体',9,0,0,'','','',0,3,
0,31,'kdtext','',0,
0,'','',2,0)


--万能表增加字段 FAdjReason
DELETE FROM t_FieldDescription WHERE FTableID=210008 AND FFieldName='FAdjReason' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(210008,'FAdjReason','STRING','Adjustment Reason','Adjustment Reason',0,0,'Adjustment Reason','FAdjReason','Adjustment Reason','Adjustment Reason')


--序事薄增加字段‘FAdjReason’
--DELETE FROM ICChatBillTitle WHERE FTypeID=68 AND FColName='FAdjReason'
--INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
--	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
--	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
--	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
--	FFormat,FFormatType)
--VALUES(@MaxInterID,68,'Adjustment Reason$','','FAdjReason','POOrder',1,1000,1,-1,1,0,1,0,'FAdjReason','v1','',0,0,
--	'',1,0,1,'Adjustment Reason$','FAdjReason$',0,0,5,'',0,0,'',0)


--套打关系增加‘FAdjReason’
DELETE FROM GLNotecitation WHERE FTemplateID='B09' AND FCode='FAdjReason'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('B09',@MaxFID,50,'FAdjReason','Adjustment Reason$',0,0,@CtlIndex,0,'FAdjReason','','',1,'FAdjReason$','Adjustment Reason$')
END
GO

EXEC p_AddStockColumn 'FAdjReason',0,0
go

--其他出库单增加字段表头FPOSNumber "POS单号"
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=68
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='B09'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='B09'
DELETE FROM ICTemplate WHERE FID='B09' AND FFieldName='POS单号'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('B09',@CtlIndex,68,'POS单号:','POS單號:','POS Number:',0,-1,1,0,0,'FPOSNumber',
2895,720,2580,480,0,0,0,
'宋体',9,0,0,'','FPOSNumber','',0,3,
0,31,'kdtext','',0,
'','','',0,0)

--万能表增加字段 FPOSNumber
--DELETE FROM t_FieldDescription WHERE FTableID=210008 AND FFieldName='FPOSNumber' 
--INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
--	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
--VALUES(210008,'FPOSNumber','STRING','POS单号','POS单号',0,0,'POS單號','POS Number','POS單號','POS Number')

--序事薄增加字段‘FPOSNumber’
DELETE FROM ICChatBillTitle WHERE FTypeID=68 AND FColName='FPOSNumber'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,68,'POS单号$','','FPOSNumber','ICStockBill',1,1000,1,-1,1,0,1,0,'FPOSNumber','v1','',0,0,
	'',1,0,1,'POS單號$','POS Number$',0,0,5,'',0,0,'',0)

--套打关系增加‘FPOSNumber’
DELETE FROM GLNotecitation WHERE FTemplateID='B09' AND FCode='FPOSNumber'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('B09',@MaxFID,50,'FPOSNumber','POS Number$',0,0,@CtlIndex,0,'FPOSNumber','','',1,'FPOSNumber$','POS Number$')
END
GO

--******************************  其他入库单  **************************************************
--其他入库单增加字段表头FAdjReason "Adjustment Reason"
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=57
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='A97'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='A97'

--表头增加字段“Adjustment Reason”'FAdjReason'
DELETE FROM ICTemplate WHERE FID='A97' AND FFieldName='FAdjReason'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('A97',@CtlIndex,57,'调整原因:','調整原因:','Adjustment Reason:',2,11400,1,1,1,'FAdjReason',
2895,720,2580,480,48,0,0,
'宋体',9,0,0,'','','',0,3,
0,31,'kdtext','',0,
0,'','',2,0)


--万能表增加字段 FAdjReason
DELETE FROM t_FieldDescription WHERE FTableID=210008 AND FFieldName='FAdjReason' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(210008,'FAdjReason','STRING','Adjustment Reason','Adjustment Reason',0,0,'Adjustment Reason','FAdjReason','Adjustment Reason','Adjustment Reason')


--序事薄增加字段‘FAdjReason’
--DELETE FROM ICChatBillTitle WHERE FTypeID=57 AND FColName='FAdjReason'
--INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
--	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
--	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
--	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
--	FFormat,FFormatType)
--VALUES(@MaxInterID,57,'Adjustment Reason$','','FAdjReason','POOrder',1,1000,1,-1,1,0,1,0,'FAdjReason','v1','',0,0,
--	'',1,0,1,'Adjustment Reason$','FAdjReason$',0,0,5,'',0,0,'',0)


--套打关系增加‘FAdjReason’
DELETE FROM GLNotecitation WHERE FTemplateID='A97' AND FCode='FAdjReason'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('A97',@MaxFID,50,'FAdjReason','Adjustment Reason$',0,0,@CtlIndex,0,'FAdjReason','','',1,'FAdjReason$','Adjustment Reason$')
END
GO

--其他入库单增加字段表头FPOSNumber "POS单号"
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=57
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='A97'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='A97'
DELETE FROM ICTemplate WHERE FID='A97' AND FFieldName='POS单号'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,
FMaxValue,FMinValue,FLookUpType,FSysMustInputItem)
Values('A97',@CtlIndex,57,'POS单号:','POS單號:','POS Number:',0,-1,1,0,0,'FPOSNumber',
2895,720,2580,480,0,0,0,
'宋体',9,0,0,'','FPOSNumber','',0,3,
0,31,'kdtext','',0,
'','','',0,0)

--万能表增加字段 FPOSNumber
DELETE FROM t_FieldDescription WHERE FTableID=210008 AND FFieldName='FPOSNumber' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(210008,'FPOSNumber','STRING','POS单号','POS单号',0,0,'POS單號','POS Number','POS單號','POS Number')

--序事薄增加字段‘FPOSNumber’
DELETE FROM ICChatBillTitle WHERE FTypeID=57 AND FColName='FPOSNumber'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,57,'POS单号$','','FPOSNumber','ICStockBill',1,1000,1,-1,1,0,1,0,'FPOSNumber','v1','',0,0,
	'',1,0,1,'POS單號$','POS Number$',0,0,5,'',0,0,'',0)

--套打关系增加‘FPOSNumber’
DELETE FROM GLNotecitation WHERE FTemplateID='A97' AND FCode='FPOSNumber'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('A97',@MaxFID,50,'FPOSNumber','POS Number$',0,0,@CtlIndex,0,'FPOSNumber','','',1,'FPOSNumber$','POS Number$')
END
GO

--******************************  调拨单  **************************************************
--调拨单增加字段表头FPOSNumber
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=75
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='D01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='D01'

--表头增加字段“POS Number”'FPOSNumber'
DELETE FROM ICTemplate WHERE FID='D01' AND FFieldName='FPOSNumber'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,FMaxValue,FMinValue,FLookUpType,FSysMustInputItem,FAllowCopy)
Values('D01',@CtlIndex,75,'POS调拨单号:','POS Number:','POS Number:',0,-1,1,0,0,'FPOSNumber',
4185,885,2580,480,0,0,0,
'宋体',9,0,0,'','','',0,3,
0,31,'kdtext','',0,
'','','',0,0,0)

--万能表增加字段 FPOSNumber
DELETE FROM t_FieldDescription WHERE FTableID=210008 AND FFieldName='FPOSNumber' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(210008,'FPOSNumber','STRING','POS调拨单号','POS Number',0,0,'POS Number','FPOSNumber','POS Number','POS Number')

--序事薄增加字段‘FPOSNumber’
DELETE FROM ICChatBillTitle WHERE FTypeID=75 AND FColName='FPOSNumber'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,75,'POS调拨单号$','','FPOSNumber','ICStockBill',1,1000,1,-1,1,0,1,0,'FPOSNumber','v1','',0,0,
	'',1,0,1,'POS Number$','POS Number$',0,0,5,'',0,0,'',0)


--套打关系增加‘FPOSNumber’
DELETE FROM GLNotecitation WHERE FTemplateID='D01' AND FCode='FPOSNumber'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('D01',@MaxFID,37,'FPOSNumber','POS调拨单号$',0,0,@CtlIndex,0,'FPOSNumber','','',1,'FPOSNumber$','POS Number$')

END
GO

--调拨单增加字段表头FSourStore
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=75
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='D01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='D01'

--表头增加字段“Origin Store”'FSourStore'
DELETE FROM ICTemplate WHERE FID='D01' AND FFieldName='FSourStore'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,FMaxValue,FMinValue,FLookUpType,FSysMustInputItem,FAllowCopy)
Values('D01',@CtlIndex,75,'发货门店:','Origin Store:','Origin Store:',0,-1,1,0,0,'FSourStore',
4185,885,2580,480,0,0,0,
'宋体',9,0,0,'','','',0,3,
0,31,'kdtext','',0,
'','','',0,0,0)

--万能表增加字段 FSourStore
DELETE FROM t_FieldDescription WHERE FTableID=210008 AND FFieldName='FSourStore' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(210008,'FSourStore','STRING','发货门店','Origin Store',0,0,'Origin Store','FSourStore','Origin Store','Origin Store')

--序事薄增加字段‘FSourStore’
DELETE FROM ICChatBillTitle WHERE FTypeID=75 AND FColName='FSourStore'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,75,'发货门店$','','FSourStore','ICStockBill',1,1000,1,-1,1,0,1,0,'FSourStore','v1','',0,0,
	'',1,0,1,'Origin Store$','Origin Store$',0,0,5,'',0,0,'',0)


--套打关系增加‘FSourStore’
DELETE FROM GLNotecitation WHERE FTemplateID='D01' AND FCode='FSourStore'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('D01',@MaxFID,37,'FSourStore','发货门店$',0,0,@CtlIndex,0,'FSourStore','','',1,'FSourStore$','Origin Store$')

END
GO

EXEC p_AddStockColumn 'FSourStore',1,0
go

--调拨单增加字段表头FDestStore
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=75
SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='D01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='D01'

--表头增加字段“Destination Store”'FDestStore'
DELETE FROM ICTemplate WHERE FID='D01' AND FFieldName='FDestStore'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,FMaxValue,FMinValue,FLookUpType,FSysMustInputItem,FAllowCopy)
Values('D01',@CtlIndex,75,'收货门店:','Destination Store:','Destination Store:',0,-1,1,0,0,'FDestStore',
4185,885,2580,480,0,0,0,
'宋体',9,0,0,'','','',0,3,
0,31,'kdtext','',0,
'','','',0,0,0)

--万能表增加字段 FDestStore
DELETE FROM t_FieldDescription WHERE FTableID=210008 AND FFieldName='FDestStore' 
INSERT INTO t_FieldDescription(FTableID,FFieldName,FFieldType,FDescription,FFieldNote,FType,FSefDefSign,
	FDescription_Cht,FDescription_En,FFieldNote_Cht,FFieldNote_En)
VALUES(210008,'FDestStore','STRING','收货门店','Destination Store',0,0,'Destination Store','FDestStore','Destination Store','Destination Store')

--序事薄增加字段‘FDestStore’
DELETE FROM ICChatBillTitle WHERE FTypeID=75 AND FColName='FDestStore'
INSERT INTO ICChatBillTitle(FInterID,FTypeID,FColCaption,FHeadSecond,FColName,FTableName,FColType,
	FColWidth,FVisible,FItemClassID,FVisForQuest,FReturnDataType,FCountPriceType,FCtlIndex,FName,
	FTableAlias,FAction,FNeedCount,FIsPrimary,FLogicAction,FStatistical,FMergeable,FVisForOrder,
	FColCaption_CHT,FColCaption_EN,FControl,FMode,FControlType,FPrecisionField,FAlign,FEditable,
	FFormat,FFormatType)
VALUES(@MaxInterID,75,'收货门店$','','FDestStore','ICStockBill',1,1000,1,-1,1,0,1,0,'FDestStore','v1','',0,0,
	'',1,0,1,'Destination Store$','Destination Store$',0,0,5,'',0,0,'',0)


--套打关系增加‘FDestStore’
DELETE FROM GLNotecitation WHERE FTemplateID='D01' AND FCode='FDestStore'
INSERT INTO GLNotecitation(FTemplateID,FID,FNoteTypeID,FCode,FCitationName,FInEntry,FIsMoney,FCtlIndex,
	FIsSum,FKeyFieldName,FExtFieldName,FTableName,FRelationID,FCitationName_EN,FCitationName_CHT)
VALUES('D01',@MaxFID,37,'FDestStore','收货门店$',0,0,@CtlIndex,0,'FDestStore','','',1,'FDestStore$','Destination Store$')

END
GO

EXEC p_AddStockColumn 'FDestStore',1,0
go

--******************************  盘盈/盘亏单  **************************************************
--盘盈单增加字段表头FIsExport
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

--SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=81
--SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='C01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='C01'

--表头增加字段“Is Export”'FIsExport'
DELETE FROM ICTemplate WHERE FID='C01' AND FFieldName='FIsExported'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,FMaxValue,FMinValue,FLookUpType,FSysMustInputItem,FAllowCopy)
Values('C01',@CtlIndex,81,'导出标识','導出標識','Export Flag',33,-1,1,3,0,'FIsExported',
4185,885,2580,480,48,0,0,
'宋体',9,0,0,'','','',0,3,
0,0,'kdtext','',0,
'0','','',0,0,1)

END
GO

--盘亏单增加字段表头FIsExport
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

--SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=81
--SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='C01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='C02'

--表头增加字段“Is Export”'FIsExport'
DELETE FROM ICTemplate WHERE FID='C02' AND FFieldName='FIsExported'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,FMaxValue,FMinValue,FLookUpType,FSysMustInputItem,FAllowCopy)
Values('C02',@CtlIndex,81,'导出标识','導出標識','Export Flag',33,-1,1,3,0,'FIsExported',
4185,885,2580,480,48,0,0,
'宋体',9,0,0,'','','',0,3,
0,0,'kdtext','',0,
'0','','',0,0,1)

END
GO

EXEC p_AddStockColumn 'FIsExported',0,0
go


--盘盈单增加字段表头FIsExportedToHQ
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

--SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=81
--SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='C01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='C01'

--表头增加字段“Is Export”'FIsExport'
DELETE FROM ICTemplate WHERE FID='C01' AND FFieldName='FIsExportedToHQ'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,FMaxValue,FMinValue,FLookUpType,FSysMustInputItem,FAllowCopy)
Values('C01',@CtlIndex,81,'HQ导出标识','HQ導出標識','HQ Export Flag',33,-1,1,3,0,'FIsExportedToHQ',
4185,885,2580,480,48,0,0,
'宋体',9,0,0,'','','',0,3,
0,0,'kdtext','',0,
'0','','',0,0,1)

END
GO

--盘亏单增加字段表头FIsExportedToHQ
BEGIN
DECLARE @MaxInterID INT,@MaxFID INT,@CtlIndex INT

--SELECT @MaxInterID=MAX(FInterID)+1 FROM ICChatBillTitle WHERE FTypeID=81
--SELECT @MaxFID=MAX(FID)+1 FROM GLNotecitation WHERE FTemplateID='C01'
SELECT @CtlIndex=MAX(FCtlIndex)+1  FROM ICTemplate WHERE FID='C02'

--表头增加字段“Is Export”'FIsExport'
DELETE FROM ICTemplate WHERE FID='C02' AND FFieldName='FIsExportedToHQ'
INSERT INTO 
ICTemplate(FID,FCtlIndex,FTabIndex,FCaption,FCaption_CHT,FCaption_EN,
FCtlType,
FLookUpCls,
FNeedSave,
FValueType,FSaveValue,FFieldName,
FLeft,FTop,FWidth,FHeight,FEnable,FRelateOutTbl,FPrint,
FFontName,FFontSize,FSelBill,FMustInput,FFilter,FRelationID,FAction,FLockA,FROB,
FDefaultCtl,FVisForBillType,FVBACtlType,FFormat,FInEntryForSPrint,
FDefaultValue,FMaxValue,FMinValue,FLookUpType,FSysMustInputItem,FAllowCopy)
Values('C02',@CtlIndex,81,'HQ导出标识','HQ導出標識','HQ Export Flag',33,-1,1,3,0,'FIsExportedToHQ',
4185,885,2580,480,48,0,0,
'宋体',9,0,0,'','','',0,3,
0,0,'kdtext','',0,
'0','','',0,0,1)

END
GO

EXEC p_AddStockColumn 'FIsExportedToHQ',0,0
go