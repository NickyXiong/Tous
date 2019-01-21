Delete from t_DataFlowDetailFunc Where FDetailFuncID = 201607
Insert into t_dataflowdetailfunc(FDetailFuncID,FFuncName,FFuncName_CHT,FFuncName_EN,FSubFuncID,FIndex,FClassName,FClassParam,FIsNormal,FHelpCode,FVisible,FAcctType,FFuncType,FEnable,FShowName,FShowName_CHT,FShowName_EN,FIsEdit,FShowSysType,FUrl,FUrlType,FFuncType_Ex) 
  Values(201607,N'调拨通知单Excel导入',N'調撥通知單Excel導入',N'Import STN by Excel',2106,7,N'Tous_Importation.Application',N'',1,N'',1,N'',0,1,Null,Null,Null,1,1,N'',N'newtab',N'base')

GO
Update t_DataFlowTimeStamp set fname = fname 
go

delete from t_ThirdPartyComponent where FTypeDetailID=71 and FComponentName ='Tous_Importation.clsOldBillsControl'
insert t_ThirdPartyComponent values(0,71,10086,'Tous_Importation.clsOldBillsControl','','')
go