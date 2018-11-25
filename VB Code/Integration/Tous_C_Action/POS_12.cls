VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POS_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements ITask
'Private notDataCLs As Object '��д������
Private datas As TB_Context.TYPE_ACTION_DATA
Private m_Dsn As String
Private m_K3Connection As ADODB.Connection
Private Function ITask_Init(k3svr As TB_Context.TYPE_K3SERVER, data As TB_Context.TYPE_ACTION_DATA, param As TB_Context.TBParameters) As Boolean
On Error GoTo HERROR
    GetDsn k3svr
    datas = data
    'Dim notDataCLs As Noteset
    'notDataCLs.SaveData "sd"
    ITask_Init = True
    Exit Function
HERROR:
    ITask_Init = False
    TB_Runtime.Log data.TaskNumber, Err.Description
End Function

Private Function ITask_Run(param As TB_Context.TBParameters) As Boolean
Dim retInfo As String '���ؽ����Ϣ
Dim rowData() As String '������
Dim billData As Variant '��������
Dim notDataCLs As Noteset '��д������
Dim obj As Object
Dim rsTemp As ADODB.Recordset
Dim strSQL As String
On Error GoTo HERROR
        Set notDataCLs = New Noteset
        With notDataCLs
            Set obj = CreateObject("Tous_M_Action.POS_12")
            billData = obj.ExportData(m_Dsn, retInfo)
            Set obj = Nothing
            
            If TB_Runtime.UboundEx(billData) > -1 Then
                notDataCLs.SaveData datas.CacheRoot & datas.FileName, billData  '��������
                TB_Runtime.K3Log datas, "Exportation is successful!", False
            End If
        End With
        Set notDataCLs = Nothing
'    End If
'    Set rsTemp = Nothing
    ITask_Run = True
    Exit Function
HERROR:
    ITask_Run = False
    TB_Runtime.K3Log datas, "Tous_C_Action.POS_12.ITask_Run" & Err.Description, True
End Function

Private Sub ITask_Dispose(param As TB_Context.TBParameters)
    'Set notDataCLs = Nothing
End Sub

Public Function GetDsn(k3svr As TB_Context.TYPE_K3SERVER) As String '����ַ���
    m_Dsn = "Persist Security Info=True;Provider=SQLOLEDB.1;User ID=" & k3svr.DBUsername & ";Password=" & k3svr.DBPassword & ";Data Source=" & k3svr.DBServer & ";Initial Catalog=" & k3svr.DBName
    'k3DsnStr = "ConnectString={" & gConString & "};UserName=" & Set_Biller & ";UserID=" & Set_BillerID & ";DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=gl;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=" & Set_BillerUUID
    GetDsn = m_Dsn
    
End Function

