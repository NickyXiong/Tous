VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "Action_SyncStockTake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Procedure : SyncPurchaseReceipt
' DateTime  : 2018-9-5
' Author    :
' Purpose   : 生成外购入库单并审核
'---------------------------------------------------------------------------------------
'
Public Function UpdateStockTakeResult(ByVal sDsn As String, strTablename As String, ByRef strInterIDs As String, ByRef sBillNOs As String, ByRef sErr As String) As Boolean
 
    Dim CNN As ADODB.Connection
    Dim bRet As Boolean
    Dim lRow As Long
    Dim i, j As Long
    Dim strSQL As String
    
On Error GoTo SynBillA01_Error
    
    '测试用
'    sDsn = "ConnectString={Persist Security Info=True;Provider=SQLOLEDB.1;User ID=sa;Password=sa;Data Source=.;Initial Catalog=Bluesea};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=BAD160C5-B4DB-4965-A8A7-0D373C885AE6"
'    strTablename = "t_POS_pur"
    
    InitDataEnv sDsn
    Set CNN = datasource.Connection

    '更新盘点数据
    strSQL = "update t1 set FQtyAct=t5.fqty,FCheckQty=t5.fqty,FAuxQtyAct=t5.fqty,FAuxCheckQty=t5.fqty,t1.FMinus=t1.FQty-t5.fqty"
    strSQL = strSQL & vbCrLf & "from ICInvBackup t1 inner join t_ICItem t2 on t1.FItemID=t2.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_Stock t3 on t1.FStockID=t3.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_StockPlace t4 on t1.FStockPlaceID=t4.FSPID"
    strSQL = strSQL & vbCrLf & "inner join " & strTablename & " t5 on t1.FBillNo=t5.FSTKScheme and t2.FNumber=t5.FSKU and t3.FNumber=t5.FStoreCode "
'    strSQL = strSQL & vbCrLf & "Where IsNull(FIsExported, 0) = 1"

    ExecSQL strSQL, sDsn
            
    Set CNN = Nothing
    UpdateStockTakeResult = True
    GetObjectContext.SetComplete

    Exit Function

SynBillA01_Error:
    UpdateStockTakeResult = False
'    GetObjectContext.SetAbort
    Set CNN = Nothing
    sErr = "Generate Purchase Receipt failed：" & sErr & Err.Description
    
End Function

Private Function ExecSQL(ByVal ssql As String, ByVal dsn As String) As ADODB.Recordset
    Dim OBJ As Object
    Dim rs As ADODB.Recordset

    Set OBJ = CreateObject("BillDataAccess.GetData")
    Set rs = OBJ.ExecuteSQL(dsn, ssql)
    Set OBJ = Nothing
    Set ExecSQL = rs
    Set rs = Nothing
End Function


