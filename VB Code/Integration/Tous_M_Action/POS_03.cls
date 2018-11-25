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
    
    Dim i As Integer
    
    Dim strSQL As String
    Dim StrSqlTemp As String
    
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
            StrSqlTemp = "SELECT 1 From ICInvBackup INNER JOIN ICStockBill ON ICInvBackup.FBillInterID=ICStockBill.FInterID"
            StrSqlTemp = StrSqlTemp & vbCrLf & "WHERE ICInvBackup.FBillNo = '" & CStr(varDatas(i)(1)) & "' and ICStockBill.FTranType in (40,43)"
            Set rs = CNN.Execute(StrSqlTemp)
            If rs.RecordCount > 0 Then
                sErr = "STK Number[" & CStr(varDatas(i)(1)) & "] has already generated Stock Take Bill in ERP system."
                GoTo UpdateStockTake_Error
            End If
        End If
           
        
        strSQL = strSQL & vbCrLf & "update t1 set FQtyAct=" & Val(varDatas(i)(4)) & ",FCheckQty=" & Val(varDatas(i)(4)) & ",FAuxQtyAct=" & Val(varDatas(i)(4)) & ","
        strSQL = strSQL & vbCrLf & "FAuxCheckQty=" & Val(varDatas(i)(4)) & ",FMinus=" & Val(varDatas(i)(4)) & "-FQty,FNote='" & CStr(varDatas(i)(2)) & "'"
        strSQL = strSQL & vbCrLf & "From ICInvBackup t1 inner join t_ICItem t2 on t1.FItemID=t2.FItemID"
        strSQL = strSQL & vbCrLf & "inner join t_Stock t3 on t1.FStockID=t3.fitemid"
        strSQL = strSQL & vbCrLf & "where t1.FBillNo='" & CStr(varDatas(i)(1)) & "' "
        strSQL = strSQL & vbCrLf & "and t3.FNumber='" & CStr(varDatas(i)(0)) & "' "  '门店代码
        strSQL = strSQL & vbCrLf & "and t2.FNumber='" & CStr(varDatas(i)(3)) & "'" 'SKU
    Next
            
    CNN.Execute strSQL
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







