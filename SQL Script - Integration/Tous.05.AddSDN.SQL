

/**Begin 新单据模板数据脚本，名称 调拨通知单    Script Date: 2018/11/25 **/


/****** Object:Data   单据描述：ICClassType    Script Date: 2018/11/25 ******/

Delete ICClassType WHERE FID=200000101
GO
INSERT INTO ICClassType(FID,FName_CHS,FName_CHT,FName_EN,FShowIndex,FTableName,FShowType,FTemplateID,FImgID,FModel,FLogic,FBillWidth,FBillHeight,FMenuControl,FFunctionID,FFilter,FBillTypeID,FIsManageBillNo,FBillNoKey,FEntryCount,FLayerCount,FLayerNames,FPrimaryKey,FEPrimaryKey,FClassTypeKey,FIndexKey,FObjectType,FObjectID,FGroupIDView,FGroupIDManage,FComponentExt,FBillNoManageType,FAccessoryTypeID,FControl,FTimeStamp,FExtBaseDataAccess,fpropertyext,FPGUID,FEGUID,FMobileBillSource,FLinkedID,FMobileIcon) 
VALUES (200000101,'调拨通知单','调拨通知单','Stock Transfer Notice',0,'t_Tous_SDN',0,200000101,'',2,3,11568,5244,'',21,'',3,0,'',0,1,'CHS=;CHT=;EN=','FID','FEntryID','FClassTypeID','FIndex',4100,200000101,2101,2102,'',1,621000265,1075775923,NULL,'','','FPGUID','FEGUID',0,0,'') 
GO

/****** Object:Data   单据分录描述：ICClassTypeEntry    Script Date: 2018/11/25 ******/

Delete ICClassTypeEntry WHERE FParentID=200000101
GO
INSERT INTO ICClassTypeEntry(FIndex,FParentID,FTableName,FLeft,FTop,FWidth,FHeight,FLayer,FEntryType,FTabIndex,FMustInput,FKeyField,FDescription_CHS,FDescription_CHT,FDescription_EN,FFilter,FUserDefine,FContainer,FCanSort,FParentEntry) 
VALUES (1,200000101,'t_Tous_SDN',0,2190,4000,4740,0,0,0,1,'','','','','',1,'',1,1) 
GO

INSERT INTO ICClassTypeEntry(FIndex,FParentID,FTableName,FLeft,FTop,FWidth,FHeight,FLayer,FEntryType,FTabIndex,FMustInput,FKeyField,FDescription_CHS,FDescription_CHT,FDescription_EN,FFilter,FUserDefine,FContainer,FCanSort,FParentEntry) 
VALUES (2,200000101,'t_Tous_SDNEntry',162,1680,11160,3360,0,1,39,0,'','','','','',1,'',1,1) 
GO

/****** Object:Data   单据模板字段描述：ICClassTableInfo    Script Date: 2018/11/25 ******/

Delete ICClassTableInfo WHERE FClassTypeID=200000101
GO
INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'内码','內碼','ISN','FID','FID','t_Tous_SDN','',1,2,0,1,1,0,3,'',0,0,'','','','','','',1,'','',56,500,255,4,'','','',1,'','PRIMARY',240,120,405,2535,'-1',5,24,0,0,0,0,'',8193,'',0,'','',0,0,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'事务类型','事務類型','Tran.Type','FClassTypeID','FClassTypeID','t_Tous_SDN','',4,2,0,1,1,0,1,'',0,0,'','FID','ICClassType','','FName','FName',1,'','',56,500,255,4,'','','',1,'','CLASSTYPEID',240,120,405,2535,'-1',11,24,0,0,0,0,'',8193,'',0,'','',0,0,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'单据编号','單據編號','Bill No','FBillNo','FBillNo','t_Tous_SDN','',1,2,2047,0,1,1,1,'',0,0,'','','','','','',1,'','',231,500,255,255,'','','',1,'','BILLNO',196,672,375,2505,'0',19,24,0,0,0,0,'',8193,'',3,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'审核流程状态','審核流程狀態','MultiCheck Flow Status','FMultiCheckStatus','FMultiCheckStatus','t_Tous_SDN','',1,2,270,0,1,0,8,'',0,0,'VALUELIST{64=不通过,16=已审核,8=待提交,4=审核中,2=未启动};VALUELIST_CHT{64=不通過,16=已審核,8=待提交,4=審核中,2=未啟動};VALUELIST_EN{64=UnPassed,16=Approved,8=To-Be Submitted,4=Approval In-Progress,2=Not Started}','','','','','',1,'','',231,500,255,255,'','2','',1,'','MCFLOWSTATUS',4851,3699,480,2270,'5',24,0,0,0,0,0,'',24579,'',9,'','',1000021,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'日期','日期','Date','FDate','FDate','t_Tous_SDN','',1,2,2047,0,1,1,0,'',0,0,'','','','','','',1,'','',61,0,8,8,'','','',1,'','',3032,672,375,2505,'2,13',25,24,0,3,23,0,'',24577,'',0,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'调出门店','调出门店','Stock-Out Store','FSCStockID','FSCStockID','t_Tous_SDN','',1,2,2047,0,1,1,1,'',1,5,'','FItemID','t_Stock','t_Stock2','FName','FNumber',1,'','0',56,500,255,4,'','','',1,'','',184,1122,375,2505,'0',26,24,0,0,10,0,'',24593,'',1,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'调出仓位','Fro调出仓位mBin','Stock-Out Bin','FSCSPID','FSCSPID','t_Tous_SDN','',1,2,2047,0,1,0,1,'',14,15,'','FSPID','t_StockPlace','','FName','FNumber',1,'FSPGroupID = (SELECT FSPGroupID FROM t_Stock WHERE FItemID = GetFldValue(FSCStockID,1))','0',56,500,255,4,'','','',1,'','',3236,2550,375,2505,'0',27,24,0,0,10,0,'',24593,'',1,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'调入门店','调入门店','Stock-In Store','FDCStockID','FDCStockID','t_Tous_SDN','',1,2,2047,0,1,1,1,'',1,5,'','FItemID','t_Stock','t_Stock1','FName','FNumber',1,'','0',56,500,255,4,'','','',1,'','',3032,1122,375,2505,'0',28,24,0,0,10,0,'',24593,'',1,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'调入仓位','调入仓位','Stock-In Bin','FDCSPID','FDCSPID','t_Tous_SDN','',1,2,2047,0,1,0,1,'',14,15,'','FSPID','t_StockPlace','t_StockPlace1','FName','FNumber',1,'FSPGroupID = (SELECT FSPGroupID FROM t_Stock WHERE FItemID = GetFldValue(FDCStockID,1))','0',56,500,255,4,'','','',1,'','',7612,2694,375,2505,'0',29,24,0,0,10,0,'',24593,'',1,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'调出单据','调出单据','Stock-Out Bill','FStockOutBill','FStockOutBill','t_Tous_SDN','',1,2,2047,0,1,0,1,'',0,0,'','','','','FName','',1,'','',231,500,50,50,'','','',1,'','',5916,672,375,2505,'0,13',31,-1,0,0,0,0,'',24595,'',0,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'调入单据','调入单据','Stock-In Bill','FStockInBill','FStockInBill','t_Tous_SDN','',1,2,2047,0,1,0,1,'',0,0,'','','','','FName','',1,'','',231,500,50,50,'','','',1,'','',8740,672,375,2505,'0,13',32,-1,0,0,0,0,'',24595,'',0,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'导出标识','导出标识','Exported Flag','FCheckBox','FIsExported','t_Tous_SDN','',1,2,12,0,1,0,4,'',0,0,'','','','','','',1,'','',104,501,255,1,'','','',1,'','',9096,84,480,2012,'7',33,-1,0,0,1,0,'',24593,'',11,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'制单人','制單人','Prep. by','FBiller','FBiller','t_Tous_SDN','',1,2,2047,-1,1,0,1,'',7,9,'','FUserID','t_User','t_User1','FName','FName',1,'','',56,500,255,4,'','','',1,'','USER',5916,1122,375,2505,'9,14,15',34,-1,0,0,6,0,'',24577,'',8,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,1,'审核人','审核人','Checker','FCheckerID','FCheckerID','t_Tous_SDN','',1,2,2047,1,1,0,1,'',7,9,'','FUserID','t_User','','FName','FName',1,'','',56,500,255,4,'','0','',1,'','',8740,1122,375,2505,'9,14,15',40,-1,0,0,6,0,'',24577,'',4,'','',1000000,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,2,'分录内码','分錄內碼','Entry ISN','FEntryID2','FEntryID','t_Tous_SDNEntry','',103,0,0,1,0,0,3,'',0,0,'','','','','','',1,'','',56,500,255,4,'','','',1,'','ENTRYKEY',240,120,350,1800,'-1',-1,24,0,0,0,0,'',8193,'',0,'','',0,0,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,2,'单据内码','單據內碼','Doc ISN','FID2','FID','t_Tous_SDNEntry','',104,0,0,1,1,1,3,'',0,0,'','','','','','',1,'','',56,500,255,4,'','','',1,'','PARENTID',240,120,350,1800,'-1',-1,24,0,0,0,0,'',8193,'',0,'','',0,0,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,2,'行号','行號','Row No.','FIndex2','FIndex','t_Tous_SDNEntry','',105,0,0,1,1,0,3,'',0,0,'','','','','','',1,'','',56,500,255,4,'','','',1,'','INDEX',240,120,350,1800,'-1',-1,-1,0,0,0,0,'',8193,'',0,'','',0,0,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,2,'商品代码','商品代码','SKU','FItemID','FItemID','t_Tous_SDNEntry','',1,2,2047,0,1,1,1,'',1,4,'','FItemID','t_ICItem','','FNumber','FNumber',1,'','0',56,500,255,4,'VerifyEntryUniquely{FItemID}','','TakeBaseData{FBaseProperty,FBaseProperty1}',1,'','',5158,7738,346,2500,'0',3,24,0,0,10,0,'',24593,'',1,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,2,'商品名称','商品名称','Description','FBaseProperty','FItemID','t_Tous_SDNEntry','',1,0,2047,0,0,0,1,'',0,4,'','FItemID','t_ICItem','','FName','',3,'','',167,500,255,0,'','','',1,'','',0,0,0,2500,'0',4,-1,0,0,0,0,'',24631,'',2,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,2,'调拨数量','调拨数量','Transfer Qty','FTransferQty','FTransferQty','t_Tous_SDNEntry','',1,2,2047,0,1,0,2,'',0,0,'Min=0 | Max=9999999999999.9999999999','','','','','',1,'','',106,9,20,13,'Vertify{FExpQty>0,CHS=Expected Qty Can not be 0@CHT=Expected Qty Can not be 0@EN=Expected Qty Can not be 0@,0}','0','',1,'','',0,0,350,2500,'1,13',5,24,1,28,1,0,'',24593,'',0,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,2,'调出数量','调出数量','Stock-Out Qty','FStockOutQty','FStockOutQty','t_Tous_SDNEntry','',1,2,2047,0,1,0,2,'',0,0,'Min=0 | Max=9999999999999.9999999999','','','','','',1,'','',106,9,20,13,'','','',1,'','',0,0,350,2500,'1,13',7,-1,1,28,1,0,'',24593,'',0,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,2,'调入数量','调入数量','Stock-In Qty','FStockInQty','FStockInQty','t_Tous_SDNEntry','',1,2,2047,0,1,0,2,'',0,0,'Min=0 | Max=9999999999999.9999999999','','','','','',1,'','',106,9,20,13,'','','',1,'','',0,0,350,2500,'1,13',10,-1,1,28,1,0,'',24593,'',0,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,2,'备注','備註','Remarks','FNOTE','FNOTE','t_Tous_SDNEntry','',1,2,2047,0,1,0,999,'',0,0,'','','','','','',0,'','',231,500,50,50,'','','',1,'','',0,0,0,2500,'0,13',10,24,0,0,0,0,'',24631,'',0,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

INSERT INTO ICClassTableInfo(FClassTypeID,FPage,FCaption_CHS,FCaption_CHT,FCaption_EN,FKey,FFieldName,FTableName,FTableNameAs,FListIndex,FListClassName,FVisible,FEnable,FNeedSave,FMustInput,FCtlType,FProperty,FLookUpType,FLookUpClassID,FLookUpList,FSRCFieldName,FSRCTableName,FSRCTableNameAs,FDSPFieldName,FFNDFieldName,FValueLocation,FFilter,FFilterGroup,FValueType,FDspColType,FEditlen,FValuePrecision,FSaveRule,FDefValue,FAction,FUserDefine,FNote,FKeyWord,FLeft,FTop,FHeight,FWidth,FCondition,FTabIndex,FLock,FSum,FPrec,FScale,FLayer,FLoadAction,FUnControl,FFont,FSourceType,FSubKey,FParentKey,FConditionExt,FFrameBorder,FFrameBorderColor,FLabelWidth,FLabelColor,FTextColor,FIsF7,FContainer,FStyle,FDescription,FDescription_CHT,FDescription_EN,IsSupportImport,FMargin,FMobileFieldLinkID,FIsShowTitle,FBackgroundImage) 
VALUES (200000101,2,'调拨差异','调拨差异','Qty Variation','FQtyVariation','FQtyVariation','t_Tous_SDNEntry','',1,2,2047,0,1,0,2,'',0,0,'Min=-9999999999999.9999999999 | Max=9999999999999.9999999999','','','','','',1,'','',106,9,20,13,'','','',1,'','',0,0,350,2500,'1,13',11,-1,1,28,1,0,'Calculate{FDecimal1=FPickQty-FDecimal}',24593,'',0,'','',0,2,0,0,0,0,0,'',0,'','','',0,'0,0,0,0',0,1,'') 
GO

/****** Object:Data   单据模板通用控件描述：ICClassCtl    Script Date: 2018/11/25 ******/

Delete ICClassCtl WHERE FClassTypeID=200000101
GO
/****** Object:Data   单据多布局列表：ICClassLayout    Script Date: 2018/11/25 ******/

Delete ICClassLayout WHERE FClassTypeID = 200000101
GO
/****** Object:Data   单据多布局明细：ICClassTableInfoLayout    Script Date: 2018/11/25 ******/

Delete ICClassTableInfoLayout WHERE FClassTypeID = 200000101
GO
/****** Object:Data   单据多布局通用控件明细：ICClassCtlLayout    Script Date: 2018/11/25 ******/

Delete ICClassCtlLayout WHERE FClassTypeID = 200000101
GO
/****** Object:Data   单据多布局的Page明细：ICClassBillEntryLayout    Script Date: 2018/11/25 ******/

Delete ICClassBillEntryLayout WHERE FClassTypeID = 200000101
GO
/****** Object:Data   单据模板单据布局控件描述：ICClassEntryEditorLayout    Script Date: 2018/11/25 ******/

Delete ICClassEntryEditorLayout WHERE FClassTypeID = 200000101 and FLayoutID = 0
GO
/****** Object:Data   序时簿综合查询设置主关系表：icCLassUnionQuery    Script Date: 2018/11/25 ******/

DELETE icCLassUnionQuery WHERE FSourClassTypeID=200000101
GO
/****** Object:Data   序时簿综合查询设置配置表：icCLassUnionQueryEntry    Script Date: 2018/11/25 ******/

DELETE icCLassUnionQueryEntry WHERE FSourClassTypeID=200000101
GO
/****** Object:Data   单据编号：ICBillNo    Script Date: 2018/11/25 ******/

 IF NOT EXISTS (SELECT * FROM  ICBillNo WHERE FBillID=200000101) 
 BEGIN 
  INSERT INTO ICBillNo(FBillID,FBillName,FPreLetter,FSufLetter,FCurNo,FBillName_CHT,FBillName_EN,FFormat,FPos,FCanAlterBillNo,FCheckAfterSave,FUseBillCodeRule,FDesc,FSingleApproval,FSingleUnApproval,FEditSelfBill) 
  VALUES (200000101,'调拨通知单_BOS','','',1,'调拨通知单_BOS','Stock Transfer Notice_BOS','00000000',200000101,0,0,1,'00000001',2,2,2) 
end 

GO

/****** Object:Data   单据编号：t_BillCodeRule    Script Date: 2018/11/25 ******/

 IF NOT EXISTS (SELECT * FROM  t_BillCodeRule WHERE FBillTypeID=200000101) 
 BEGIN 
INSERT INTO t_BillCodeRule(FBillTypeID,FClassIndex,FProjectID,FProjectVal,FFormatIndex,FLength,FAddChar,FReChar,FBillType,FIsBy) 
VALUES ('200000101',2,3,'1',0,8,'','','200000101',0) 

 END 

/****** Object:Data   单据权限类弄表：t_ObjectAccessType    Script Date: 2018/11/25 ******/

Delete t_ObjectAccessType WHERE FObjectType=4100 And FObjectID =200000101
GO
INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,1,2097152,1048576,'查看','查看','查看','View','查看','View') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,2,1024,0,'查看凭证','查看凭证','查看憑證','View Voucher','查看憑證','View Voucher') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,3,131072,3145728,'新增','新增','新增','New','新增','New') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,4,65536,3145728,'删除','删除','刪除','Delete','刪除','Delete') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,5,4194304,3145728,'修改','修改','修改','Edit','修改','Edit') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,6,16,3145728,'引出内部数据','引出内部数据','引出內部資料','Export Internal Data','引出內部資料','Export Internal Data') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,7,16384,0,'导出单据模板','导出单据模板','導出單據模板','Export Doc Template','導出單據模板','Export Doc Template') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,8,32768,3145728,'打印','打印','列印','Print','列印','Print') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,9,128,0,'按单生成凭证','按单生成凭证','按單生成憑證','Generate Voucher by Doc','按單生成憑證','Generate Voucher by Doc') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,10,512,0,'合并生成凭证','合并生成凭证','合併生成憑證','Consolidate and Generate Voucher','合併生成憑證','Consolidate and Generate Voucher') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,11,2048,0,'发送邮件','发送邮件','發送郵件','Send Mail','發送郵件','Send Mail') 
GO

INSERT INTO t_ObjectAccessType(FObjectType,FObjectID,FIndex,FAccessMask,FAccessUse,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,12,8192,0,'发送短信','发送短信','發送短信','Send SMS','發送短信','Send SMS') 
GO

/****** Object:Data   单据权限对象表：t_ObjectType    Script Date: 2018/11/25 ******/

Delete t_ObjectType WHERE FObjectType=4100 And FObjectID =200000101
GO
INSERT INTO t_ObjectType(FObjectType,FObjectID,FName,FDescription,FName_cht,FName_en,FDescription_cht,FDescription_en) 
VALUES (4100,200000101,'调拨通知单','','调拨通知单','Stock Transfer Notice',NULL,NULL) 
GO

/****** Object:Data   单据权限表：t_ObjectAccess    Script Date: 2018/11/25 ******/

Delete t_ObjectAccess WHERE FObjectType=4100 And FObjectID =200000101
GO
INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,0) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,1) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,2) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,3) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,4) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,5) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,6) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,7) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,8) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,9) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,10) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,11) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2102,4100,200000101,12) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2101,4100,200000101,0) 
GO

INSERT INTO t_ObjectAccess(FGroupID,FObjectType,FObjectID,FIndex) 
VALUES (2101,4100,200000101,1) 
GO

/****** Object:Data   选单关联：ICClassLink    Script Date: 2018/11/25 ******/

Delete ICClassLink WHERE FDestClassTypeID = 200000101 OR FSourClassTypeID =  200000101
GO
/****** Object:Data   选单关联明细：ICClassLinkEntry    Script Date: 2018/11/25 ******/

Delete ICClassLinkEntry WHERE FDestClassTypeID=200000101 OR FSourClassTypeID =  200000101
GO
/****** Object:Data   选单钩稽明细：ICClassLinkCommit    Script Date: 2018/11/25 ******/

 DELETE ICClassLinkCommit WHERE FDstClsTypID = 200000101 OR FSrcClsTypID =  200000101
GO
/****** Object:Data   选单流程基本信息：ICClassWorkFlow    Script Date: 2018/11/25 ******/


GO
DECLARE @MaxID AS INT

SELECT @MaxID = ISNUll(Max(FID),10000) FROM ICClassWorkFlow


UPDATE ICMaxNum SET FMaxNum = @MaxID + 1 WHERE FTableName = 'ICClassWorkFlow'
GO


/****** Object:Data   选单流程关联单据：ICClassWorkFlowBill    Script Date: 2018/11/25 ******/

DELETE ICClassWorkFlowBill WHERE FClassTypeID = 200000101
GO
DELETE ICClassWorkFlowBill WHERE FID < 0 AND FEntryID NOT IN ( SELECT MAX(FEntryID)  FROM ICClassWorkFlowBill WHERE FID < 0 GROUP BY FClassTypeID HAVING COUNT(FClassTypeID) >= 1 ) 
DELETE ICClassWorkFlowBill WHERE FID=( SELECT MAX(FID) FROM ICClassWorkFlow )  and FClassTypeID in (SELECT FClassTypeID from ICCLASSWORKFLOWBILL where FID<0)
 UPDATE ICClassWorkFlowBill SET FID =( SELECT MAX(FID) FROM ICClassWorkFlow ) WHERE FID < 0 
 DELETE ICClassWorkFlowBill WHERE FID < 0 
/****** Object:Data   选单流程关联关系：ICClassWorkFlowJoin    Script Date: 2018/11/25 ******/

 DELETE ICClassWorkFlowJoin WHERE FDestClassTypeID = 200000101 OR FSourClassTypeID = 200000101
GO
 UPDATE ICClassWorkFlowJoin SET FID =( SELECT MAX(FID) FROM ICClassWorkFlow ) WHERE FID < 0 
 DELETE ICClassWorkFlowJoin WHERE FID < 0 
/****** Object:Data   网络控制模板数据：ICClassMutex    Script Date: 2018/11/25 ******/

Delete ICClassMutex WHERE FClassTypeID=200000101
GO
/****** Object:Data   网络控制互斥表明细数据：t_Mutex    Script Date: 2018/11/25 ******/

DELETE t_Mutex 
WHERE FFuncId IN
(Select FFuncID from icClassMutex where FClassTypeID=200000101)
OR FForbidden IN
(Select FFuncID from icClassMutex where FClassTypeID=200000101)

GO
/****** Object:Data   网络控制操作定义明细：t_SysFunction    Script Date: 2018/11/25 ******/

DELETE t_SysFunction 
WHERE FFuncId IN
(Select FFuncID from icClassMutex where FClassTypeID=200000101)

GO
/****** Object:Data   打印控制模板数据：ICPrintMaxCount    Script Date: 2018/11/25 ******/

Delete ICPrintMaxCount WHERE FID =200000101
GO
INSERT INTO ICPrintMaxCount(FPos,FID,FName,FName_CHT,FName_EN,FPrintControlOrNo,FPrintMaxCount,FPrintAfterCheck,FSaveByUser) 
VALUES (200000101,200000101,'调拨通知单_BOS','调拨通知单_BOS','Stock Transfer Notice_BOS',0,0,-1,-1) 
GO

/****** Object:Data   多级审核配置模板数据：ICClassMCFlowInfo    Script Date: 2018/11/25 ******/

DELETE ICClassMCFlowInfo WHERE FID = 200000101
GO
/****** Object:Data   多级审核的工作流任务（各审核级次）模板数据：ICClassMCTasks    Script Date: 2018/11/25 ******/

DELETE ICClassMCTasks WHERE FID = 200000101
GO
/****** Object:Data   多级审核的流程流转和规则模板数据：ICClassMCRule    Script Date: 2018/11/25 ******/

DELETE ICClassMCRule WHERE FID = 200000101
GO
/****** Object:Data   单据操作明细模板数据：ICClassBillAction    Script Date: 2018/11/25 ******/

DELETE ICClassBillAction WHERE FClassTypeID = 200000101
GO
DECLARE @MaxID AS INT

SELECT @MaxID = ISNUll(Max(FID),10000) FROM ICClassBillAction

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 1,200000101,'mnuBOSViewVoucher','查看凭证','查看憑證','View Voucher','查看凭证','查看憑證','View Voucher',0,1024,0,7,1,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 2,200000101,'mnuEditCheckRoad','查看审核路线','查看審核路線','View Approval Routing','查看审核路线','查看審核路線','View Approval Routing',997,0,0,7,0,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 3,200000101,'mnuEditCopy','复制','複製','Copy','复制','複製','Copy',39,131072,0,7,1,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 4,200000101,'mnuEditCopyAllRow','复制所有行','複製所有行','Copy All Rows','复制所有行','複製所有行','Copy All Rows',39,0,0,7,0,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 5,200000101,'mnuEditCopyDoc','按单复制','按單複製','Copy by Doc','按单复制','按單複製','Copy by Doc',64,131072,0,7,1,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 6,200000101,'mnuEditCopyEntry','按分录合并复制','按分錄合併複製','Merge Selected Entries and Copy','按分录合并复制','按分錄合併複製','Merge Selected Entries and Copy',64,131072,0,7,1,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 7,200000101,'mnuEditDelete','删除','刪除','Delete','删除','刪除','Delete',64,65536,0,7,1,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 8,200000101,'mnuEditModify','修改','修改','Edit','修改','修改','Edit',4160,4194304,0,7,1,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 9,200000101,'mnuEditMultiCheck','多级审核','多級審核','Multi-level Approval','多级审核','多級審核','Multi-level Approval',997,0,0,7,0,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 10,200000101,'mnuEditStartMultiCheck','启动多级审核','啟動多級審核','Enable Multi-level Approval','启动多级审核','啟動多級審核','Enable Multi-level Approval',997,0,0,7,0,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 11,200000101,'mnuEditUnMultiCheck','驳回审核','駁回審核','Reject Approval','驳回审核','駁回審核','Reject Approval',997,0,0,7,0,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 12,200000101,'mnuEditUnPassMultiCheck','审核不通过','審核不通過','Failed the Check','审核不通过','審核不通過','Failed the Check',997,0,0,7,0,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 13,200000101,'mnuExportData','引出内部数据','引出內部資料','Export Internal Data','引出内部数据','引出內部資料','Export Internal Data',5056,16,0,7,1,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 14,200000101,'mnuFileExportBillTpl','导出单据模板','導出單據模板','Export Doc Template','导出单据模板','導出單據模板','Export Doc Template',0,16384,0,7,1,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 15,200000101,'mnuFileImportBill','导入单据数据','導入單據資料','Import Doc Data','导入单据数据','導入單據資料','Import Doc Data',0,16384,0,7,1,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 16,200000101,'mnuFileNew','新增','新增','New','新增','新增','New',103,131072,0,7,1,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 17,200000101,'mnuFilePrint','打印','列印','Print','打印','列印','Print',5095,32768,0,7,1,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 18,200000101,'mnuFileSave','保存','保存','Save','保存','保存','Save',6,0,0,7,0,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 19,200000101,'mnuFileView','查看','查看','View','查看','查看','View',960,2097152,0,7,1,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 20,200000101,'mnuMakeOneToOne','按单生成凭证','按單生成憑證','Generate Voucher by Doc','按单生成凭证','按單生成憑證','Generate Voucher by Doc',0,128,0,7,1,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 21,200000101,'mnuManualSubmitMultiCheck','提交审核','提交審核','Submit for Approval','提交审核','提交審核','Submit for Approval',997,0,0,7,0,0,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 22,200000101,'mnuMultiToOne','合并生成凭证','合併生成憑證','Consolidate and Generate Voucher','合并生成凭证','合併生成憑證','Consolidate and Generate Voucher',0,512,0,7,1,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 23,200000101,'mnuPriceInfoQuery','价格资料查询','價格資料查詢','Query Price','价格资料查询','價格資料查詢','Query Price',0,16384,0,7,1,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 24,200000101,'mnuSendEmail','发送邮件','發送郵件','Send Mail','发送邮件','發送郵件','Send Mail',0,2048,0,7,1,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 25,200000101,'mnuSendIMMessage','发送云之家消息','發送雲之家消息','Send CloudHub Message','发送云之家消息','發送雲之家消息','Send CloudHub Message',0,0,0,7,0,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 26,200000101,'mnuSendMessage','发送消息','發送消息','Send Message','发送消息','發送消息','Send Message',0,0,0,7,0,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 27,200000101,'mnuSendSMS','发送短信','發送短信','Send SMS','发送短信','發送短信','Send SMS',0,8192,0,7,1,1,1,0,0,1,0) 

INSERT INTO ICClassBillAction(FID,FClassTypeID,FCode,FName_CHS,FName_CHT,FName_EN,FDescriptions_CHS,FDescriptions_CHT,FDescriptions_EN,FEnvironment,FRightMask,FRightUseMask,FRightGroupMask,FRightControl,FAddLog,FIsPredefine,FIsFirst,FShortCut,FListActionCirculation,FOperationType) 
VALUES ( @MaxID + 28,200000101,'mnuStockQuery','库存查询','庫存查詢','Query Inventory','库存查询','庫存查詢','Query Inventory',0,0,0,7,0,1,1,0,0,1,0) 


UPDATE ICMaxNum SET FMaxNum = @MaxID + 29 WHERE FTableName = 'ICClassBillAction'
GO


/****** Object:Data   单据操作日志模板数据：t_LogFunction    Script Date: 2018/11/25 ******/

DELETE t_LogFunction WHERE FFunctionID LIKE 'BOS200000101%'
GO
INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuEditDelete','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuFileExportBillTpl','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuFileImportBill','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuFileSave','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuMakeOneToOne','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuMultiToOne','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuPriceInfoQuery','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuSendEmail','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuSendIMMessage','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuSendMessage','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuSendSMS','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

INSERT INTO t_LogFunction(FFunctionID,FFunctionName,FSubSysID,FFunctionName_Cht,FFunctionName_EN) 
VALUES ('BOS200000101_mnuStockQuery','仓存管理系统-调拨通知单',21,'倉存管理系統-调拨通知单','Inventory Module-Stock Transfer Notice') 
GO

/****** Object:Data   单据操作位置模板数据：ICClassActionPosition    Script Date: 2018/11/25 ******/

DELETE ICClassActionPosition WHERE FClassTypeID = 200000101
GO
/****** Object:Data   单据扩展服务模板数据：ICClassActionList    Script Date: 2018/11/25 ******/

DELETE ICClassActionList WHERE FClassTypeID = 200000101
GO
INSERT INTO ICClassActionList(FClassTypeID,FClassActionID,FObject,FDefineType,FSourceType,FSourceField,FAction,FExpression,FOrder,FDescription) 
VALUES (200000101,1,'FItemID',0,4096,'FAction','FBASEPROPERTY,FBASEPROPERTY2,FBaseProperty1','',1,'值更新事件') 
GO

INSERT INTO ICClassActionList(FClassTypeID,FClassActionID,FObject,FDefineType,FSourceType,FSourceField,FAction,FExpression,FOrder,FDescription) 
VALUES (200000101,21,'FItemID',0,4096,'FSaveRule','FItemID','',1,'保存规则') 
GO

INSERT INTO ICClassActionList(FClassTypeID,FClassActionID,FObject,FDefineType,FSourceType,FSourceField,FAction,FExpression,FOrder,FDescription) 
VALUES (200000101,22,'EVENT',0,8192,'FAfterLoadBillAction','','@IsMultiCheckStarted=1',1,'Added by BOS Create') 
GO



/**End  新单据模板数据脚本，名称 调拨通知单    Script Date: 2018/11/25 **/




--Create Data Table
--ERP库存备份
IF NOT EXISTS(SELECT 1 FROM SysObjects WHERE NAME = 't_Tous_SDN' AND XTYPE = 'U')
CREATE TABLE [dbo].[t_Tous_SDN](
	[FID] [int] NOT NULL ,
	[FClassTypeID] [int] NOT NULL DEFAULT ((0)),
	[FBillNo] [nvarchar](255) NOT NULL DEFAULT (''),
	[FMultiCheckStatus] [nvarchar](255) NOT NULL DEFAULT (''),
	[FDate] [datetime] NULL,
	[FSCStockID] [int] NOT NULL DEFAULT ((0)),
	[FSCSPID] [int] NOT NULL DEFAULT ((0)),
	[FDCStockID] [int] NOT NULL DEFAULT ((0)),
	[FDCSPID] [int] NOT NULL DEFAULT ((0)),
	[FStockOutBill] [nvarchar](50) NOT NULL,
	[FStockInBill] [nvarchar](50) NOT NULL,
	[FIsExported] [bit] NOT NULL DEFAULT ((0)),
	[FCheckerID] [int] NOT NULL DEFAULT ((0)),
	[FBiller] [int] NOT NULL DEFAULT ((0)),
PRIMARY KEY CLUSTERED 
(
	[FID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


IF NOT EXISTS(SELECT 1 FROM SysObjects WHERE NAME = 't_Tous_SDNEntry' AND XTYPE = 'U')
CREATE TABLE [dbo].[t_Tous_SDNEntry](
	[FEntryID] [int] IDENTITY(1,1) NOT NULL ,
	[FID] [int] NOT NULL,
	[FIndex] [int] NOT NULL DEFAULT ((0)),
	[FItemID] [int] NOT NULL DEFAULT ((0)),
	[FTransferQty] [decimal](23, 10) NOT NULL DEFAULT ((0)),
	[FStockOutQty] [decimal](23, 10) NOT NULL DEFAULT ((0)),
	[FStockInQty] [decimal](23, 10) NOT NULL DEFAULT ((0)),
	[FQtyVariation] [decimal](23, 10) NOT NULL DEFAULT ((0)),
	[FNOTE] [nvarchar](50) NOT NULL DEFAULT ('')
) ON [PRIMARY]

GO