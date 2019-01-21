VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************
'Function: 更新实盘数据 (POS TO ERP)
'Date:
'*******************************************************

Option Explicit

Private datasource As CDataSource
Private m_oParse As CParse
Private strLanguage As String

Private Function InitDataEnv(ByVal sToParse As String) As Boolean
    Set m_oParse = New CParse
    If m_oParse.ParseString(sToParse) Then
        Set datasource = New CDataSource
        Set datasource.ParseObject = m_oParse
    Else
        Err.Raise EBS_E_TypeMismatch, "ParseString"
    End If
    strLanguage = datasource.ParseObject.LanguageID
End Function

'---------------------------------------------------------------------------------------
' Procedure : UpdateStockTake
' DateTime  :
' Author    :
' Purpose   : 生成外购入库单并审核 (POS TO ERP)
'---------------------------------------------------------------------------------------
'
Public Function UpdateStockTake(ByVal sDsn As String, varDatas As Variant, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String) As Boolean
    Dim CNN As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim rsPOInfo As ADODB.Recordset
    
    Dim lFItemID As Long, lUnitID As Long, lFStockID As Long
    
    Dim i As Long
    
    Dim strSql As String
    Dim strSQLtemp As String
    
On Error GoTo UpdateStockTake_Error

    InitDataEnv sDsn
    Set CNN = datasource.Connection

    For i = LBound(varDatas) To UBound(varDatas)
        If i = 0 Then
            '检查盘点方案是否存在，不存在则不允许重复导入
            Set rs = CNN.Execute("select 1 from icstockcheckprocess where FProcessID='" & CStr(varDatas(i)(1)) & "'")
            If rs.RecordCount <= 0 Then
                sErr = "STK Number[" & CStr(varDatas(i)(1)) & "] has NOT existed in ERP system."
                GoTo UpdateStockTake_Error
            End If
            
            '检查是否存在，不存在则不允许重复导入
            strSQLtemp = "SELECT 1 From ICInvBackup t1 INNER JOIN ICStockBill t2 ON t1.FBillInterID=t2.FInterID"
            strSQLtemp = strSQLtemp & vbCrLf & "inner join icstockcheckprocess t3 on t1.FInterID=t3.FID"
            strSQLtemp = strSQLtemp & vbCrLf & "WHERE t3.FProcessID = '" & CStr(varDatas(i)(1)) & "' and t2.FTranType in (40,43)"
            Set rs = CNN.Execute(strSQLtemp)
            If rs.RecordCount > 0 Then
                sErr = "STK Number[" & CStr(varDatas(i)(1)) & "] has already generated Stock Take Bill in ERP system."
                GoTo UpdateStockTake_Error
            End If
        End If
            
        
        strSQLtemp = "select 1 From ICInvBackup t1 inner join t_ICItem t2 on t1.FItemID=t2.FItemID"
        strSQLtemp = strSQLtemp & vbCrLf & "inner join t_Stock t3 on t1.FStockID=t3.fitemid"
        strSQLtemp = strSQLtemp & vbCrLf & "inner join icstockcheckprocess t4 on t1.FInterID=t4.FID"
        strSQLtemp = strSQLtemp & vbCrLf & "where t4.FProcessID='" & CStr(varDatas(i)(1)) & "' "
        strSQLtemp = strSQLtemp & vbCrLf & "and t3.FNumber='" & CStr(varDatas(i)(0)) & "' "  '门店代码
        strSQLtemp = strSQLtemp & vbCrLf & "and t2.FNumber='" & CStr(varDatas(i)(3)) & "'" 'SKU
        Set rs = CNN.Execute(strSQLtemp)
        
        If rs.RecordCount <= 0 Then
'            sErr = sErr & "Store[" & CStr(varDatas(i)(0)) & "]/SKU[" & CStr(varDatas(i)(3)) & "] has NOT existed in stock take report."
            
            '如果不在盘点方案中，则增加记录
             '检查物料是否存在
             lFItemID = GetItemIDByNumber(CNN, CStr(varDatas(i)(3)), 4, lUnitID)
             If lFItemID = 0 Then
                 sErr = sErr & "Material Master Data [" & CStr(varDatas(i)(3)) & "] is NOT created in ERP system."
                 GoTo UpdateStockTake_Error
             End If
            
            '检查仓库是否存在
            lFStockID = GetItemID(CNN, CStr(varDatas(i)(0)), 5)
            If lFStockID = 0 Then
                 sErr = sErr & "Warehouse Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                 GoTo UpdateStockTake_Error
            End If
            
            strSql = strSql & vbCrLf & "insert ICInvBackup "
            strSql = strSql & vbCrLf & "select top 1 t1.FBrNo," & lFItemID & ",FBatchNo,FStockID,0,FBal,FDateBackup," & Val(varDatas(i)(4)) & ",0,FDateCheck,t1.FChecker,FCheckerMan," & Val(varDatas(i)(4)) & ","
            strSql = strSql & vbCrLf & "FMinusPercent,'" & CStr(varDatas(i)(2)) & "',FSelect,FAdj," & Val(varDatas(i)(4)) & ",FBillInterID,t1.FMaxBillInterID,FStockPlaceID,t1.FKFPeriod,FKFDate,"
            strSql = strSql & vbCrLf & "FUnitID,0," & Val(varDatas(i)(4)) & "," & Val(varDatas(i)(4)) & ",t1.FSecUnitID,FInterID,FSecQty,FAdjQty,FAuxPropID,FMTONo,FSecQtyAct,"
            strSql = strSql & vbCrLf & "FSecCheckQty,FSecAdjQty,FSecMinus,FCheckerID,t1.FStatus,FMultiCheckLevel1,FMultiCheckLevel2,FMultiCheckLevel3,"
            strSql = strSql & vbCrLf & "FMultiCheckLevel4,FMultiCheckLevel5,FMultiCheckLevel6,FTranType,FBillNo,t1.FDate,FCancelLation,FBillerID,FMultiCheckDate1,"
            strSql = strSql & vbCrLf & "FMultiCheckDate2 , FMultiCheckDate3, FMultiCheckDate4, FMultiCheckDate5, FMultiCheckDate6, FCurCheckLevel, FCheckDate, FSupplyID"
            strSql = strSql & vbCrLf & "from ICInvBackup t1 inner join icstockcheckprocess t4 on t1.FInterID=t4.FID"
            strSql = strSql & vbCrLf & "where t4.FProcessID='" & CStr(varDatas(i)(1)) & "' "
            
        Else
            strSql = strSql & vbCrLf & "update t1 set FQtyAct=" & Val(varDatas(i)(4)) & ",FCheckQty=" & Val(varDatas(i)(4)) & ",FAuxQtyAct=" & Val(varDatas(i)(4)) & ","
            strSql = strSql & vbCrLf & "FAuxCheckQty=" & Val(varDatas(i)(4)) & ",FMinus=" & Val(varDatas(i)(4)) & "-FQty,FNote='" & CStr(varDatas(i)(2)) & "'"
            strSql = strSql & vbCrLf & "From ICInvBackup t1 inner join t_ICItem t2 on t1.FItemID=t2.FItemID"
            strSql = strSql & vbCrLf & "inner join t_Stock t3 on t1.FStockID=t3.fitemid"
            strSql = strSql & vbCrLf & "inner join icstockcheckprocess t4 on t1.FInterID=t4.FID"
            strSql = strSql & vbCrLf & "where t4.FProcessID='" & CStr(varDatas(i)(1)) & "' "
            strSql = strSql & vbCrLf & "and t3.FNumber='" & CStr(varDatas(i)(0)) & "' "  '门店代码
            strSql = strSql & vbCrLf & "and t2.FNumber='" & CStr(varDatas(i)(3)) & "'" 'SKU
        End If
        
    Next
            
    CNN.Execute strSql
    Set CNN = Nothing
    UpdateStockTake = True
    GetObjectContext.SetComplete

   Exit Function

UpdateStockTake_Error:
    UpdateStockTake = False
    sErr = "Update Stock Take ：" & sErr & Err.Description
    GetObjectContext.SetAbort
    Set CNN = Nothing
    

End Function







