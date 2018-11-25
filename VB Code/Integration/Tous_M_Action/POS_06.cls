VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_06"
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
Private m_dataFeil(0 To 14) As String

Private Sub Class_Initialize() '初始化,导出字段位置说明
    m_dataFeil(0) = "FBillNo" 'Style Number
    m_dataFeil(1) = "FNoticeNo" 'SKU #
    m_dataFeil(2) = "FPurchaseType" 'Import Currency
    m_dataFeil(3) = "FDate" 'Import Price
    m_dataFeil(4) = "FWHCode" 'Retail Currency
    m_dataFeil(5) = "FWHBin" 'Retail Price
    m_dataFeil(6) = "FStoreCode" 'Retail Price
    m_dataFeil(7) = "FStoreBin" 'Retail Price
    m_dataFeil(8) = "FBillerID" 'Retail Price
    m_dataFeil(9) = "FBillDate" 'Retail Price
    m_dataFeil(10) = "FNote" 'Retail Price
    m_dataFeil(11) = "FItemNumber" 'Retail Price
    m_dataFeil(12) = "FQty" 'Retail Price
    m_dataFeil(13) = "FHQRetailPrice" 'Retail Price
    m_dataFeil(14) = "FDicount" 'Retail Price
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
Public Function ExportData(ByVal m_Dsn As String, ByRef retInfo As String, ByVal lInterID As Long) As Variant '更新物料
Dim rowData() As String
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim m_billData As Variant
Dim n As Integer, i As Integer

On Error GoTo HERROR
    m_Dsns = m_Dsn
    Init '初始化
    strSQL = "select FHQNumber FBillNo,''FNoticeNo,isnull(t4.FName,'') FPurchaseType,CONVERT(varchar(100),t1.FDate, 23) FDate,t3.FNumber FWHCode,'000' FWHBin,"
    strSQL = strSQL & vbCrLf & "t3.FNumber FStoreCode,'000' FStoreBin,t3.FNumber FBillerID,CONVERT(varchar(100),t1.FDate, 23) FBillDate,"
    strSQL = strSQL & vbCrLf & "t2.FNote,t5.FNumber FItemNumber,Convert(decimal(18,2),t2.FQty)FQty,isnull(t5.FHQRetailPrice,0)FHQRetailPrice,'1'FDicount"
    strSQL = strSQL & vbCrLf & "from POOrder t1 inner join POOrderEntry t2 on t1.FInterID=t2.FInterID"
    strSQL = strSQL & vbCrLf & "inner join t_Department t3 on t1.FDeptID=t3.FItemID"
    strSQL = strSQL & vbCrLf & "left join t_SubMessage t4 on t1.FPurchaseType=t4.FInterID"
    strSQL = strSQL & vbCrLf & "inner join t_ICItem t5 on t2.FItemID=t5.FItemID"
    strSQL = strSQL & vbCrLf & "where t1.FInterID=" & CStr(lInterID)

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
    
    strSQL = "update POOrder set FIsExported=1 where FInterID=" & CStr(lInterID)
    m_MDConnection.Execute strSQL
    
    Set rs = Nothing
    Exit Function
HERROR:
    retInfo = "Tous_M_Action.POS_06.ExportData:" & Err.Description
    '写日志
End Function





