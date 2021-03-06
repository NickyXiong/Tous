VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "HQ_04"
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
Private m_dataFeil(0 To 10) As String

Private Sub Class_Initialize() '初始化,导出字段位置说明
    m_dataFeil(0) = "FBillNo" 'Style Number
    m_dataFeil(1) = "FDate" 'Import Price
    m_dataFeil(2) = "FStoreCode" 'Retail Price
    m_dataFeil(3) = "FItemNumber" 'Retail Price
    m_dataFeil(4) = "FQty" 'Retail Price
    m_dataFeil(5) = "FTime" 'Retail Price
    m_dataFeil(6) = "FROB" 'Retail Price
    m_dataFeil(7) = "FDiscountAmount" 'Retail Price
    m_dataFeil(8) = "FDiscountRate" 'Retail Price
    m_dataFeil(9) = "FConsignAmount" 'Retail Price
    m_dataFeil(10) = "FConsigneeAdd" 'Retail Price
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
Public Function ExportData(ByVal m_Dsn As String, ByRef retInfo As String, ByVal strPOSNo As String) As Variant '更新物料
Dim rowData() As String
Dim strSql As String
Dim rs As ADODB.Recordset
Dim m_billData As Variant
Dim n As Integer, i As Integer

On Error GoTo HError
    m_Dsns = m_Dsn
    Init '初始化
    strSql = "select cast(t1.FBillNo as integer) FBillNo,CONVERT(varchar(100),t1.FDate, 23) FDate, t3.FNumber FStoreCode,t5.FNumber FItemNumber,"
    strSql = strSql & vbCrLf & "t2.FQty FQty,t1.FExplanation FTime,t1.FROB,FDiscountAmount,case when FDiscountAmount+FConsignAmount=0 then 0 else FDiscountAmount/(FDiscountAmount+FConsignAmount)*100 end FDiscountRate,FConsignAmount,t1.FConsigneeAdd "
    strSql = strSql & vbCrLf & "from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
    strSql = strSql & vbCrLf & "inner join t_Stock t3 on t2.FDCStockID=t3.FItemID"
    strSql = strSql & vbCrLf & "inner join t_ICItem t5 on t2.FItemID=t5.FItemID"
    strSql = strSql & vbCrLf & "inner join t_MeasureUnit t4 on t5.FUnitID=t4.FItemID"
    strSql = strSql & vbCrLf & "inner join t_Organization t6 on t1.FSupplyID=t6.FItemID"
    strSql = strSql & vbCrLf & "Where FTranType = 21 and t1.FPOSNumber='" & strPOSNo & "'"
    strSql = strSql & vbCrLf & "order by t1.FInterID,t2.FEntryID "

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
    
    strSql = "update ICStockBill set FIsExportedToHQ=1 where FTranType = 21 and FPOSNumber='" & strPOSNo & "'"
    m_MDConnection.Execute strSql
    
    Set rs = Nothing
    Exit Function
HError:
    retInfo = "Tous_M_Action.HQ_04.ExportData:" & Err.Description
    '写日志
End Function


