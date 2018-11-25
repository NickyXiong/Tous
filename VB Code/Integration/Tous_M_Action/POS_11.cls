VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_11"
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
Private m_dataFeil(0 To 13) As String

Private Sub Class_Initialize() '初始化,导出字段位置说明
    m_dataFeil(0) = "FNumber"
    m_dataFeil(1) = "FEAN13"
    m_dataFeil(2) = "FName"
    m_dataFeil(3) = "FBrandCode"
    m_dataFeil(4) = "FBrandName"
    m_dataFeil(5) = "FYearCode"
    m_dataFeil(6) = "FYearName"
    m_dataFeil(7) = "FSeasonCode"
    m_dataFeil(8) = "FSeasonName"
    m_dataFeil(9) = "FCateCode"
    m_dataFeil(10) = "FCateName"
    m_dataFeil(11) = "FSubCateCode"
    m_dataFeil(12) = "FSubCateName"
    m_dataFeil(13) = "FHQRetailPrice"
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
Public Function ExportData(ByVal m_Dsn As String, ByRef retInfo As String) As Variant '更新物料
Dim rowData() As String
Dim strSql As String
Dim rs As ADODB.Recordset
Dim m_billData As Variant
Dim n As Integer, i As Integer

On Error GoTo HERROR
    m_Dsns = m_Dsn
    Init '初始化
    strSql = "select FNumber,FEAN13,FName,'Tous'FBrandCode,'Tous'FBrandName,''FYearCode,''FYearName,''FSeasonCode,''FSeasonName,"
    strSql = strSql & vbCrLf & "''FCateCode,''FCateName,''FSubCateCode,''FSubCateName,FHQRetailPrice from t_icitem where isnull(FUpdatedFlag,0)=1"
    Set rs = m_MDConnection.Execute(strSql)
    
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
    
    strSql = "update t_icitem set FUpdatedFlag=0 where isnull(FUpdatedFlag,0)=1"
    m_MDConnection.Execute strSql
    
    Set rs = Nothing
    Exit Function
HERROR:
    retInfo = "Tous_M_Action.POS_11.ExportData:" & Err.Description
    '写日志
End Function





