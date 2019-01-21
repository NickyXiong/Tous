VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************
'Function: Export SKU Master
'Date: 2013/11/08
'*******************************************************
Option Explicit

Private m_Dsns As String
Private m_K3Dsn As String
Private m_MDConnection As ADODB.Connection
Private m_dataFeil(0 To 7) As String

Private Sub Class_Initialize() '初始化,导出字段位置说明
    m_dataFeil(0) = "FBillNo" 'Style Number
    m_dataFeil(1) = "FDate" 'SKU #
    m_dataFeil(2) = "FStoreCode" 'Import Currency
    m_dataFeil(3) = "FStoreBin" 'Import Price
    m_dataFeil(4) = "FBillerID" 'Retail Currency
    m_dataFeil(5) = "FBillDate" 'Retail Price
    m_dataFeil(6) = "FNote" 'Retail Price
    m_dataFeil(7) = "FItemNumber" 'Retail Price
End Sub
Private Sub Init() '初始化
    Set m_MDConnection = New ADODB.Connection
    With m_MDConnection
        .ConnectionString = m_Dsns
        .CommandTimeout = 15
        .ConnectionTimeout = 15
        .CursorLocation = adUseClient
        .Open
    End With
    
    m_K3Dsn = "ConnectString={" & m_Dsns & "};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=cc;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;MachineName=127.0.0.1;K3Version=KUE;UUID=00000000-0000-0000-0000-000000000000"
End Sub
Public Function ExportData(ByVal m_Dsn As String, ByRef retInfo As String, ByVal strProcess As String) As Variant '更新物料
Dim rowData() As String
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim m_billData As Variant
Dim n As Long, i As Long

On Error GoTo HERROR
    m_Dsns = m_Dsn
    Init '初始化
    strSQL = "select t4.FProcessID FBillNo,CONVERT(varchar(100),t1.FDateBackup, 23)FDate,t3.FNumber FStoreCode,'000'FStoreBin,"
    strSQL = strSQL & vbCrLf & "t3.FNumber FBillerID,CONVERT(varchar(100),t1.FDateBackup, 23)FBillDate,t1.FNote ,t2.FNumber FItemNumber"
    strSQL = strSQL & vbCrLf & "From ICInvBackup t1 inner join t_ICItem t2 on t1.FItemID=t2.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_Stock t3 on t1.FStockID=t3.fitemid"
    strSQL = strSQL & vbCrLf & "inner join icstockcheckprocess t4 on t1.FInterID=t4.FID"
    strSQL = strSQL & vbCrLf & "where t4.FProcessID='" & strProcess & "'"
    strSQL = strSQL & vbCrLf & "order by t3.FNumber,t2.FNumber  "

    Set rs = m_MDConnection.Execute(strSQL)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then '导出数据
            ReDim m_billData(rs.RecordCount - 1)
            ReDim rowData(UBound(m_dataFeil))
            For n = 1 To rs.RecordCount
                'rowdata = Nothing
                For i = 0 To UBound(m_dataFeil) '6列导出数据
                    rowData(i) = CStr(CNulls(rs(m_dataFeil(i)), ""))
                Next i
                m_billData(n - 1) = rowData
                rs.MoveNext
            Next
        End If
    End If
    ExportData = m_billData
    
    strSQL = "update icstockcheckprocess set FIsExported=1 where FProcessID='" & strProcess & "'"
    m_MDConnection.Execute strSQL
    
    Set rs = Nothing
    Exit Function
HERROR:
    retInfo = "Tous_M_Action.POS_09.ExportData:" & Err.Description
    '写日志
End Function






