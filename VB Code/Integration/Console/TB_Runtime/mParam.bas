Attribute VB_Name = "mParam"
Option Explicit

Public K3Connection As ADODB.Connection
Public Tasks As TB_Context.TBTasks

Public Const CONST_DIR_SETTING As String = "Setting\"
Public Const CONST_DIR_LOG As String = "Log\"
Public Const CONST_FILE_K3SERVER As String = "K3Server.ini"
Public Const CONST_FILE_SMTP As String = "SMTP.ini"
Public Const CONST_FILE_REMOTE As String = "Remote.list"
Public Const CONST_FILE_METADATA As String = "Metadata.list"
Public Const CONST_FILE_TASK As String = "Task.list"
Public Const CONST_FILE_ACTION As String = "TaskAction.list"
Public Const CONST_FILE_EMAIL As String = "Send.list"

Public Sub InitConnection()
    Dim k3svr As TYPE_K3SERVER
    Dim o As LocalContent
    
    If K3Connection Is Nothing Then
        Set K3Connection = New ADODB.Connection
    End If
    If K3Connection.State = 0 Then
        Set o = New LocalContent
        k3svr = o.GetK3Server
        
        With K3Connection
            .ConnectionString = "Persist Security Info=True;Provider=SQLOLEDB.1;User ID=" & k3svr.DBUsername & ";Password=" & k3svr.DBPassword & ";Data Source=" & k3svr.DBServer & ";Initial Catalog=" & k3svr.DBName
            .ConnectionTimeout = 15
            .CursorLocation = adUseClient
            .Open
        End With
        Set o = Nothing
    End If
    
    If Tasks Is Nothing Then
        Set Tasks = New TB_Context.TBTasks
        Set o = New LocalContent
        Set Tasks = o.GetTasks
        Set o = Nothing
    End If
End Sub

Public Sub Dispose()
    If Not K3Connection Is Nothing Then
        If K3Connection.State = 1 Then
            K3Connection.Close
        End If
        Set K3Connection = Nothing
    End If
    Set Tasks = Nothing
End Sub
