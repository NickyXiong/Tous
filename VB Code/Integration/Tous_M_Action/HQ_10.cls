VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "HQ_10"
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
Private m_dataFeil(0 To 11) As String

Private Sub Class_Initialize() '初始化,导出字段位置说明
    m_dataFeil(0) = "FBillNo" 'Style Number
    m_dataFeil(1) = "FDate" 'Import Price
    m_dataFeil(2) = "FStoreCode" 'Retail Price
    m_dataFeil(3) = "FEntryID" 'Retail Price
    m_dataFeil(4) = "FItemNumber" 'Retail Price
    m_dataFeil(5) = "FEAN8" 'Retail Price
    m_dataFeil(6) = "FEAN13" 'Retail Price
    m_dataFeil(7) = "FQty" 'Retail Price
    m_dataFeil(8) = "FAuxPrice" 'Retail Price
    m_dataFeil(9) = "FLineAmount" 'Retail Price
    m_dataFeil(10) = "FHQCostPrice" 'Retail Price
    m_dataFeil(11) = "FSupplierCode" 'Retail Price
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
Dim strSql As String
Dim rs As ADODB.Recordset
Dim m_billData As Variant
Dim n As Integer, i As Integer

On Error GoTo HERROR
    m_Dsns = m_Dsn
    Init '初始化
'    strSql = "select t1.FBillNo,CONVERT(varchar(100),t1.FDate, 23) FDate,t3.FNumber FStoreCode,t2.FEntryID,t5.FNumber FItemNumber,t5.FEAN8,t5.FEAN13,"
'    strSql = strSql & vbCrLf & "t2.FQty,t6.FAuxPrice,t2.FQty*t6.FAuxPrice FLineAmount,t5.FHQCostPrice,t8.FNumber FSupplierCode"
'    strSql = strSql & vbCrLf & "from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
'    strSql = strSql & vbCrLf & "inner join t_Stock t3 on t2.FDCStockID=t3.FItemID"
'    strSql = strSql & vbCrLf & "inner join t_ICItem t5 on t2.FItemID=t5.FItemID"
'    strSql = strSql & vbCrLf & "inner join POOrderEntry t6 on t2.FSourceInterId=t6.FInterID and t2.FSourceEntryID=t6.FEntryID"
'    strSql = strSql & vbCrLf & "inner join POOrder t7 on t6.FInterID=t7.FInterID"
'    strSql = strSql & vbCrLf & "inner join t_Supplier t8 on t7.FSupplyID=t8.FItemID"
'    strSql = strSql & vbCrLf & "where t1.FTranType in (1,10) and FROB=1 and IsNull(FIsExportedToHQ, 0) = 0 and t1.FInterID=" & CStr(lInterID)
'    strSql = strSql & vbCrLf & "order by t1.finterid,t2.FEntryID"
    
    strSql = "select t1.FBillNo,CONVERT(varchar(100),t1.FDate, 23) FDate,t3.FNumber FStoreCode,t2.FEntryID,t5.FNumber FItemNumber,t5.FEAN8,t5.FEAN13,"
    strSql = strSql & vbCrLf & "t2.FQty,t6.FAuxPrice,t2.FQty*t6.FAuxPrice FLineAmount,t5.FHQCostPrice,t8.FNumber FSupplierCode"
    strSql = strSql & vbCrLf & "from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
    strSql = strSql & vbCrLf & "inner join t_Stock t3 on t2.FDCStockID=t3.FItemID"
    strSql = strSql & vbCrLf & "inner join t_ICItem t5 on t2.FItemID=t5.FItemID"
    strSql = strSql & vbCrLf & "left join POOrderEntry t6 on t2.FSourceInterId=t6.FInterID and t2.FSourceEntryID=t6.FEntryID"
    strSql = strSql & vbCrLf & "left join POOrder t7 on t6.FInterID=t7.FInterID"
    strSql = strSql & vbCrLf & "left join t_Supplier t8 on t7.FSupplyID=t8.FItemID"
    strSql = strSql & vbCrLf & "where t1.FTranType in (1) and FROB=1 and IsNull(FIsExportedToHQ, 0) = 0 and t1.FInterID= & CStr(lInterID)"
    strSql = strSql & vbCrLf & "Union"
    strSql = strSql & vbCrLf & "select t1.FBillNo,CONVERT(varchar(100),t1.FDate, 23) FDate,t3.FNumber FStoreCode,t2.FEntryID,t5.FNumber FItemNumber,t5.FEAN8,t5.FEAN13,"
    strSql = strSql & vbCrLf & "t2.FQty,t6.FAuxPrice,t2.FQty*t6.FAuxPrice FLineAmount,t5.FHQCostPrice,t8.FNumber FSupplierCode"
    strSql = strSql & vbCrLf & "from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
    strSql = strSql & vbCrLf & "inner join t_Stock t3 on t2.FDCStockID=t3.FItemID"
    strSql = strSql & vbCrLf & "inner join t_ICItem t5 on t2.FItemID=t5.FItemID"
    strSql = strSql & vbCrLf & "inner join POOrderEntry t6 on t2.FSourceInterId=t6.FInterID and t2.FSourceEntryID=t6.FEntryID"
    strSql = strSql & vbCrLf & "inner join POOrder t7 on t6.FInterID=t7.FInterID"
    strSql = strSql & vbCrLf & "inner join t_Supplier t8 on t7.FSupplyID=t8.FItemID"
    strSql = strSql & vbCrLf & "where t1.FTranType in (10) and FROB=1 and IsNull(FIsExportedToHQ, 0) = 0 and t1.FInterID= & CStr(lInterID)"
    strSql = strSql & vbCrLf & "order by t1.FBillNo,t2.FEntryID"
    
    
    
    
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
    
    strSql = "update ICStockBill set FIsExportedToHQ=1 where FInterID=" & CStr(lInterID)
    m_MDConnection.Execute strSql
    
    Set rs = Nothing
    Exit Function
HERROR:
    retInfo = "Tous_M_Action.HQ_10.ExportData:" & Err.Description
    '写日志
End Function


