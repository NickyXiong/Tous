VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "HQ_09"
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
Private m_dataFeil(0 To 12) As String

Private Sub Class_Initialize() '初始化,导出字段位置说明
    m_dataFeil(0) = "FBillNo" 'Style Number
    m_dataFeil(1) = "FDate" 'Import Price
    m_dataFeil(2) = "FStoreCode" 'Retail Price
    m_dataFeil(3) = "FStoreBin" 'Retail Price
    m_dataFeil(4) = "FBillerID" 'Retail Price
    m_dataFeil(5) = "FBillDate" 'Retail Price
    m_dataFeil(6) = "FNote" 'Retail Price
    m_dataFeil(7) = "FItemNumber" 'Retail Price
    m_dataFeil(8) = "FQty" 'Retail Price
    m_dataFeil(9) = "FHQRetailPrice" 'Retail Price
    m_dataFeil(10) = "FDicount" 'Retail Price
    m_dataFeil(11) = "FUnit" 'Retail Price
    m_dataFeil(12) = "FDesc" 'Retail Price
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
Public Function ExportData(ByVal m_Dsn As String, ByRef retInfo As String, ByVal lInterID As Long, ByVal strStore As String) As Variant '更新物料
Dim rowData() As String
Dim strSql As String
Dim rs As ADODB.Recordset
Dim m_billData As Variant
Dim n As Integer, i As Integer

On Error GoTo HERROR
    m_Dsns = m_Dsn
    Init '初始化
    strSql = "select t1.FBillNo,CONVERT(varchar(100),t1.FDate, 23) FDate,t3.FNumber FStoreCode,'000' FStoreBin, "
    strSql = strSql & vbCrLf & "t3.FNumber FBillerID,CONVERT(varchar(100),t1.FDate, 23) FBillDate,t2.FNote,t5.FNumber FItemNumber,"
    strSql = strSql & vbCrLf & "case when FTranType in (40,10) then Convert(decimal(18,2),t2.FQty) else Convert(decimal(18,2),t2.FQty)*-1 end FQty,"
    strSql = strSql & vbCrLf & "isnull(t5.FHQRetailPrice,0)FHQRetailPrice,'1'FDicount,t4.FNumber FUnit,t5.FName FDesc"
    strSql = strSql & vbCrLf & "from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
    strSql = strSql & vbCrLf & "inner join t_Stock t3 on t2.FDCStockID=t3.FItemID"
    strSql = strSql & vbCrLf & "inner join t_ICItem t5 on t2.FItemID=t5.FItemID"
    strSql = strSql & vbCrLf & "inner join t_MeasureUnit t4 on t5.FUnitID=t4.FItemID "
    strSql = strSql & vbCrLf & "where FTranType in (40,10,43,29) and t1.FInterID=" & CStr(lInterID) & " and t3.FNumber='" & strStore & "'"

    Set rs = m_MDConnection.Execute(strSql)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then '导出数据
            ReDim m_billData(rs.RecordCount - 1)
            ReDim rowData(UBound(m_dataFeil))
            For n = 1 To rs.RecordCount
                'rowdata = Nothing
                For i = 0 To UBound(m_dataFeil)
                    rowData(i) = CStr(CNulls(rs(m_dataFeil(i)), ""))
                Next i
                m_billData(n - 1) = rowData
                rs.MoveNext
            Next
        End If
    End If
    ExportData = m_billData
    
'    strSql = "update ICStockBill set FIsExportedToHQ=1 where FTranType in (40,10,43,29)  and FInterID=" & CStr(lInterID)
'    m_MDConnection.Execute strSql
    
    Set rs = Nothing
    Exit Function
HERROR:
    retInfo = "Tous_M_Action.HQ_09.ExportData:" & Err.Description
    '写日志
End Function
