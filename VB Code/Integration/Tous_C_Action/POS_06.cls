VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POS_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements ITask
'Private notDataCLs As Object '读写数据类
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
Dim retInfo As String '返回结果信息
Dim rowData() As String '行数据
Dim billData As Variant '单据数据
Dim notDataCLs As Noteset '读写数据类
Dim obj As Object
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim i As Integer, lInstr As Long
Dim strNewFileName As String
On Error GoTo HERROR

    strSQL = "select FInterID,FBillNo from POOrder where FCheckerID>0 and isnull(fisexported,0)=0 "
    Set rs = TB_Runtime.ExecuteSQL(strSQL)
'    Set rs = ExecSQL(strSQL, m_Dsn)
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            Sleep (2000)
            
            Set notDataCLs = New Noteset
            
            With notDataCLs
                lInstr = InStr(1, datas.FileName, ".")
                strNewFileName = Left(datas.FileName, lInstr - 7) & Format(Time, "hhmmss") & Mid(datas.FileName, lInstr, 4)
                datas.FileName = strNewFileName
                
                Set obj = CreateObject("Tous_M_Action.POS_06")
                billData = obj.ExportData(m_Dsn, retInfo, rs.Fields("FInterID").Value)
                Set obj = Nothing
                
                If TB_Runtime.UboundEx(billData) > -1 Then
                    
                    notDataCLs.SaveData datas.CacheRoot & strNewFileName, billData  '导出数据
                    
'                    TB_Runtime.K3Log datas, "Exportation is successful!", False
                    TB_Runtime.K3Log datas, "Exportation is successful!Bill No:[" & rs.Fields("FBillNo").Value & "]", False
                End If
            End With
            Set notDataCLs = Nothing
            
            rs.MoveNext
        Next
    End If
    
    ITask_Run = True
    Exit Function
HERROR:
    ITask_Run = False
    TB_Runtime.K3Log datas, "Tous_C_Action.POS_06.ITask_Run" & Err.Description, True
End Function

Private Sub ITask_Dispose(param As TB_Context.TBParameters)
    'Set notDataCLs = Nothing
End Sub

Public Function GetDsn(k3svr As TB_Context.TYPE_K3SERVER) As String '组合字符串
    m_Dsn = "Persist Security Info=True;Provider=SQLOLEDB.1;User ID=" & k3svr.DBUsername & ";Password=" & k3svr.DBPassword & ";Data Source=" & k3svr.DBServer & ";Initial Catalog=" & k3svr.DBName
    'k3DsnStr = "ConnectString={" & gConString & "};UserName=" & Set_Biller & ";UserID=" & Set_BillerID & ";DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=gl;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=" & Set_BillerUUID
    GetDsn = m_Dsn
    
End Function

