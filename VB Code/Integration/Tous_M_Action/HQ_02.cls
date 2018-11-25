VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "HQ_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Dsns As String
Private m_K3Dsn As String
Private m_MDConnection As ADODB.Connection

Private Sub Init()
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

Public Function UpdatePrice(ByVal m_Dsn As String, ByVal dic As KFO.Dictionary, ByRef retInfo As String) As Boolean   '保存物料
Dim strSQL As String
Dim rs As ADODB.Recordset
On Error GoTo HERROR
    m_Dsns = m_Dsn
    Init  '初始化
    
    strSQL = "select 1 from t_ICItem WHERE FNumber ='" & CStr(dic("SKU")) & "'"
    Set rs = m_MDConnection.Execute(strSQL)
    If rs.RecordCount <= 0 Then
        retInfo = "SKU can not be found in the system."
        UpdatePrice = False
        Exit Function
    End If
    
    If CStr(dic("Type")) = "retail_price" Then
        strSQL = "UPDATE t_ICItem SET "
        strSQL = strSQL & " FHQRetailPrice='" & Val(dic("Value")) & "', "
        strSQL = strSQL & " FUpdatedFlag =1 WHERE FNumber ='" & CStr(dic("SKU")) & "'"
        m_MDConnection.Execute strSQL
    End If
    
    If CStr(dic("Type")) = "cost" Then
        strSQL = "UPDATE t_ICItem SET "
        strSQL = strSQL & " FHQCostPrice='" & Val(dic("Value")) & "', "
        strSQL = strSQL & " FUpdatedFlag =1 WHERE FNumber ='" & CStr(dic("SKU")) & "'"
        m_MDConnection.Execute strSQL
    End If
    
    UpdatePrice = True
    
    Set rs = Nothing
    GetObjectContext.SetComplete
    Exit Function
HERROR:
    GetObjectContext.SetAbort
    UpdatePrice = False
    retInfo = Err.Description
    
End Function


Public Function ExceProc(ByVal procName As String, ByVal par1 As ADODB.Parameter, ByVal par2 As ADODB.Parameter, ByVal con As ADODB.Connection) As ADODB.Parameters
    Dim adoComm As Object
    '// 创建一个对象，我们用来调用存储过程
    Set adoComm = CreateObject("ADODB.Command")
    With adoComm
        '// 设置连接，假设 adoConn 为已经连接的 ADODB.Connection 对象
        .ActiveConnection = con
        '// 类型为存储过程，adCmdStoredProc = 4
        .CommandType = 4
        '// 存储过程名称
        .CommandText = procName
        '// 设置输入参数
        .Parameters(1) = par1
        .Parameters(2) = par2
        '// 执行存储过程
        .Execute
        Set ExceProc = .Parameters
    End With
    '// 释放对象
    Set adoComm = Nothing
End Function



