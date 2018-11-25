VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "HQ_06"
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
Private m_dataFeil(0 To 9) As String

Private Sub Class_Initialize() '��ʼ��,�����ֶ�λ��˵��
    m_dataFeil(0) = "FBillNo" 'Style Number
    m_dataFeil(1) = "FDate" 'Import Price
    m_dataFeil(2) = "FStoreCode" 'Retail Price
    m_dataFeil(3) = "FItemNumber" 'Retail Price
    m_dataFeil(4) = "FQty" 'Retail Price
    m_dataFeil(5) = "FUnit" 'Retail Price
    m_dataFeil(6) = "FDesc" 'Retail Price
    m_dataFeil(7) = "FSupplier" 'Retail Price
    m_dataFeil(8) = "FType" 'Retail Price
    m_dataFeil(9) = "FEntryID" 'Retail Price
End Sub
Private Sub Init() '��ʼ��
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
Public Function ExportData(ByVal m_Dsn As String, ByRef retInfo As String, ByVal lInterID As Long) As Variant '��������
Dim rowData() As String
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim m_billData As Variant
Dim n As Integer, i As Integer

On Error GoTo HERROR
    m_Dsns = m_Dsn
    Init '��ʼ��
    strSQL = "select t1.FBillNo,CONVERT(varchar(100),t1.FDate, 23) FDate, t3.FNumber FStoreCode,t5.FNumber FItemNumber,"
    strSQL = strSQL & vbCrLf & "(t2.FQty*t1.FROB) FQty,t4.FNumber FUnit,t5.FName FDesc,t6.FNumber FSupplier,2 FType,t2.FEntryID"
    strSQL = strSQL & vbCrLf & "from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
    strSQL = strSQL & vbCrLf & "inner join t_Stock t3 on t2.FDCStockID=t3.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_ICItem t5 on t2.FItemID=t5.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_MeasureUnit t4 on t5.FUnitID=t4.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_Supplier t6 on t1.FSupplyID=t6.FItemID"
    strSQL = strSQL & vbCrLf & "Where FTranType = 1 And FROB = -1 and t1.FInterID=" & CStr(lInterID)

    '���ӵ���������
    strSQL = strSQL & vbCrLf & "union"
    strSQL = strSQL & vbCrLf & "select t1.FPOSNumber,CONVERT(varchar(100),t1.FDate, 23) FDate, t3.FNumber FStoreCode,t5.FNumber FItemNumber,"
    strSQL = strSQL & vbCrLf & "(t2.FQty*t1.FROB) FQty,t4.FNumber FUnit,t5.FName FDesc,t1.FDestStore FSupplier,1 FType,t2.FEntryID"
    strSQL = strSQL & vbCrLf & "from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
    strSQL = strSQL & vbCrLf & "inner join t_Stock t3 on t2.FSCStockID=t3.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_ICItem t5 on t2.FItemID=t5.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_MeasureUnit t4 on t5.FUnitID=t4.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_Stock t6 on t2.FDCStockID=t6.FItemID"
    strSQL = strSQL & vbCrLf & "Where FTranType = 41 and t6.FNumber='T-WH' and t1.FInterID= " & CStr(lInterID)
    strSQL = strSQL & vbCrLf & "order by FBillNo,t2.FEntryID"
    
    Set rs = m_MDConnection.Execute(strSQL)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then '��������
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
    
    strSQL = "update ICStockBill set FIsExportedToHQ=1 where FInterID=" & CStr(lInterID)
    m_MDConnection.Execute strSQL
    
    Set rs = Nothing
    Exit Function
HERROR:
    retInfo = "Tous_M_Action.HQ_06.ExportData:" & Err.Description
    'д��־
End Function
