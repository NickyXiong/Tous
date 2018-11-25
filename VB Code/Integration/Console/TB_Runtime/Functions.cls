VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Functions"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get K3ConnectionString(k3svr As TYPE_K3SERVER) As String
    K3ConnectionString = "Persist Security Info=True;Provider=SQLOLEDB.1;User ID=" & k3svr.DBUsername & ";Password=" & k3svr.DBPassword & ";Data Source=" & k3svr.DBServer & ";Initial Catalog=" & k3svr.DBName
End Property

Public Property Get DirectoryAsSet() As String
    Dim locDir As String
    locDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    locDir = locDir & CONST_DIR_SETTING & "\"
    DirectoryAsSet = locDir
End Property

Public Property Get DirectoryAsLog(TaskNumber As String) As String
    Dim locDir As String
    locDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    locDir = locDir & CONST_DIR_LOG & TaskNumber & "\"
    DirectoryAsLog = locDir
End Property

Public Function SetUserDefine(ByVal Section As String, ByVal Key As String, ByVal Val As String)
    Dim path As String
    path = DirectoryAsSet & "UserDefine.ini"
    mINI.Config_INI_Patch = path
    mINI.SetIniValue Section, Key, Val
End Function

Public Function RetUserDefine(ByVal Section As String, ByVal Key As String, Optional ByVal optVal As String = "")
    Dim path As String
    path = DirectoryAsSet & "UserDefine.ini"
    mINI.Config_INI_Patch = path
    RetUserDefine = mINI.GetIniValue(Section, Key, optVal)
End Function

Public Sub Log(ByVal TaskNumber As String, ByVal Info As String)
    Dim lpt As Long
    Dim locDir As String
    
    locDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    locDir = locDir & CONST_DIR_LOG & TaskNumber & "\"
    MakeDirByLocal locDir
    
    locDir = locDir & Format(Date, "YYYYMMDD") & ".log"
    
    lpt = FreeFile
    Info = "[" & Format(Now, "YYYYMMDD hhmmss") & "] " & Info
    Open locDir For Append As #lpt
    Print #lpt, Info
    Close #lpt
End Sub

Public Sub K3Log(o As TYPE_ACTION_DATA, ByVal Info As String, Optional ByVal isSend As Boolean = False)
    Dim Index As Long
    Dim bRet As Boolean
    Dim TaskDesc As String
    Dim sSQL As String
On Error GoTo HERROR

    mParam.InitConnection
    Info = Replace(Info, "'", "''")
    
    sSQL = "INSERT INTO t_TB_Log(FActionID,FAction,FTask,FFilename,FLocDir,FDir,FFormat,FDate,FFrom,FTo,FFromID,FToID,FDescription,FIsSend,FTaskNumber,FActionNumber)" & vbCrLf
    sSQL = sSQL & "VALUES("
    
    If o.FromID = ENUM_FTPROOT.ROOT_K3 Then
        sSQL = sSQL & "1,'Export'"
    Else
        sSQL = sSQL & "2,'Import'"
    End If
    
    Index = mParam.Tasks.Lookup(o.TaskNumber)
    sSQL = sSQL & ",'" & mParam.Tasks.task(Index).Description & "','" & o.FileName & "','" & o.BackupRoot & "','" & o.Root & "','TXT-UTF8;Tab-delimited',CONVERT(VARCHAR(20),GETDATE(),120)"
    
    Select Case o.FromID
        Case ENUM_FTPROOT.ROOT_3PL
            sSQL = sSQL & ",'3PL'"
        Case ENUM_FTPROOT.ROOT_K3
            sSQL = sSQL & ",'K/3 ERP'"
        Case ENUM_FTPROOT.ROOT_POS
            sSQL = sSQL & ",'POS'"
        Case ENUM_FTPROOT.ROOT_TBHQ
            sSQL = sSQL & ",'TB HQ'"
    End Select
    
    Select Case o.ToID
        Case ENUM_FTPROOT.ROOT_3PL
            sSQL = sSQL & ",'3PL'"
        Case ENUM_FTPROOT.ROOT_K3
            sSQL = sSQL & ",'K/3 ERP'"
        Case ENUM_FTPROOT.ROOT_POS
            sSQL = sSQL & ",'POS'"
        Case ENUM_FTPROOT.ROOT_TBHQ
            sSQL = sSQL & ",'TB HQ'"
    End Select
    
    sSQL = sSQL & "," & o.FromID & "," & o.ToID & ",'" & Info & "'," & IIf(isSend, 1, 0) & ",'" & o.TaskNumber & "','" & o.ActionNumber & "')"
    
    mParam.K3Connection.BeginTrans
    bRet = True
    mParam.K3Connection.Execute sSQL
    mParam.K3Connection.CommitTrans
    bRet = False
    Exit Sub
HERROR:
    If bRet Then
        mParam.K3Connection.RollbackTrans
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function UboundEx(vArray As Variant) As Long
On Error GoTo HERROR
    UboundEx = UBound(vArray)
    Exit Function
HERROR:
    UboundEx = -1
End Function

Public Function MakeDirByLocal(ByVal path As String) As Boolean
    Dim lIndex As Integer
    Dim tempPath As String
On Error GoTo HERROR
    
    If Left(path, 2) = "\\" Then
        lIndex = InStr(3, path, "\")
        lIndex = InStr(lIndex + 1, path, "\")
    Else
        lIndex = InStr(1, path, "\")
    End If
    
    While lIndex > 0
        tempPath = Left(path, lIndex)
        If tempPath <> "\\" Then
            If Dir(tempPath, vbDirectory) = "" Then
                MkDir tempPath
            End If
        End If
        lIndex = InStr(lIndex + 1, path, "\")
    Wend

    MakeDirByLocal = True
    Exit Function
HERROR:
    MakeDirByLocal = False
End Function

Public Function ExecuteSQL(sSQL As String) As ADODB.Recordset
    mParam.InitConnection
    Set ExecuteSQL = mParam.K3Connection.Execute(sSQL)
End Function

Public Function ExecuteTransSQL(sSQL As String) As ADODB.Recordset
    Dim bRet As Boolean
On Error GoTo HERROR
    mParam.InitConnection
    mParam.K3Connection.BeginTrans
    bRet = True
    Set ExecuteTransSQL = mParam.K3Connection.Execute(sSQL)
    mParam.K3Connection.CommitTrans
    bRet = False
    Exit Function
HERROR:
    If bRet Then
        mParam.K3Connection.RollbackTrans
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Sub Class_Terminate()
    mParam.Dispose
End Sub
