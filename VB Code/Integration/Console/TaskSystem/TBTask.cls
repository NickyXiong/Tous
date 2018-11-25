VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements TB_Runtime.ILoad

Private ThisK3Server As TYPE_K3SERVER
Private ThisTasks As TB_Context.TBTasks
Private ThisMetadatas As TB_Context.TBMetadatas
Private ThisActions As TB_Context.TBActions
Private ThisShare As TB_Context.XZShareLock
Private ThisRemotes As TB_Context.TBRemotes
Private ThisFileSystem As FileSystemObject
Private ThisParamter As TB_Context.TBParameters

Private Function ILoad_Init(param As TB_Context.TBParameters) As Boolean
On Error GoTo HERROR
    
    mFunc.LogEx "Task:" & mStart.Parameter.Value("@TASK") & " Action:" & mStart.Parameter.Value("@ACTION") & " Initialize"
    DoEvents
    
    Dim o As TB_Runtime.LocalContent
    Set o = New TB_Runtime.LocalContent
    ThisK3Server = o.GetK3Server
    Set ThisTasks = o.GetTasks
    Set ThisMetadatas = o.GetMetadatas
    Set ThisActions = o.GetActions
    Set ThisRemotes = o.GetRemotes
    Set o = Nothing
    
    Set ThisShare = New TB_Context.XZShareLock
    Set ThisFileSystem = New FileSystemObject
    Set ThisParamter = New TB_Context.TBParameters
    
    ILoad_Init = True
    Exit Function
    
HERROR:
    mFunc.LogEx Err.Description
End Function

Private Sub ILoad_Run(param As TB_Context.TBParameters)
    mFunc.LogEx "Task:" & mStart.Parameter.Value("@TASK") & " Action:" & mStart.Parameter.Value("@ACTION") & " Run"
    DoEvents
    DoAction
End Sub

Private Sub ILoad_Dispose(param As TB_Context.TBParameters)
    mFunc.LogEx "Task:" & mStart.Parameter.Value("@TASK") & " Action:" & mStart.Parameter.Value("@ACTION") & " Dispose"
    DoEvents
    
    Set ThisMetadatas = Nothing
    Set ThisActions = Nothing
    Set ThisShare = Nothing
    Set ThisRemotes = Nothing
    Set ThisFileSystem = Nothing
    Set ThisParamter = Nothing
End Sub

Private Sub DoAction()
    Dim Index As Long
    Dim act As TYPE_ACTION
On Error GoTo HERROR

    For Index = 0 To ThisActions.Size - 1
        act = ThisActions.Action(Index)
        If mStart.Parameter.Value("@TASK") = act.TaskNumber And mStart.Parameter.Value("@ACTION") = act.ActionNumber Then
            DoActionEx act
        End If
    Next Index
    Exit Sub
HERROR:
    mFunc.LogEx Err.Description
End Sub

Private Sub DoActionEx(act As TYPE_ACTION)
    Dim Index As Long
    Dim meta As TYPE_METADATA
       
    For Index = 0 To ThisMetadatas.Size - 1
        
        If mStart.Is_CommandExit Then
            Exit For
        End If
        
        meta = ThisMetadatas.Metadata(Index)
        If act.MetadataNumber = meta.ID Then
            If meta.ToID = ENUM_FTPROOT.ROOT_K3 Then
                DoActionByMetadatas act, meta
            Else
                DoActionByMetadatas2 act, meta
            End If
        End If
        
        DoEvents
    Next Index
End Sub

Private Sub DoActionByMetadatas(act As TYPE_ACTION, meta As TYPE_METADATA)
    Dim oFolder As Folder, oFile As File
    Dim rmt As TYPE_REMOTE, data As TYPE_ACTION_DATA
    Dim locRoot As String
    
    rmt = ThisRemotes.Remote(ThisRemotes.Lookup(ENUM_FTPROOT.ROOT_3PL Or ENUM_FTPROOT.ROOT_POS Or ENUM_FTPROOT.ROOT_TBHQ))
    locRoot = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    TB_Runtime.MakeDirByLocal locRoot & rmt.CacheRoot & meta.Directory
    Set oFolder = ThisFileSystem.GetFolder(locRoot & rmt.CacheRoot & meta.Directory)
    
    If oFolder.Files.Count > 0 Then
        For Each oFile In oFolder.Files
            
            If mStart.Is_CommandExit Then
                Exit For
            End If
        
            If SameNameString(meta, oFile.Name) Then
                With data
                    .TaskNumber = "TASK" & act.TaskNumber
                    .ActionNumber = "ACTION" & act.ActionNumber
                    .MetadataNumber = act.MetadataNumber
                    .FromID = meta.FromID
                    .ToID = meta.ToID
                    .Filename = oFile.Name
                    .CacheRoot = RetLocalDirectory(locRoot & rmt.CacheRoot & meta.Directory)
                    .BackupRoot = RetLocalDirectory(locRoot & rmt.BackupRoot & meta.Directory)
                    .Root = meta.Directory
                End With
                
                DoActionByMetadata act.dllName, data
            End If
            Set oFile = Nothing
        Next
        
        DoEvents
    Else
    
    End If
    Set oFolder = Nothing
End Sub

Private Sub DoActionByMetadatas2(act As TYPE_ACTION, meta As TYPE_METADATA)
    Dim rmt As TYPE_REMOTE, data As TYPE_ACTION_DATA
    Dim locRoot As String
    
    rmt = ThisRemotes.Remote(ThisRemotes.Lookup(ENUM_FTPROOT.ROOT_3PL Or ENUM_FTPROOT.ROOT_POS Or ENUM_FTPROOT.ROOT_TBHQ))
    locRoot = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    TB_Runtime.MakeDirByLocal locRoot & rmt.CacheRoot & meta.Directory
    
    With data
        .TaskNumber = "TASK" & act.TaskNumber
        .ActionNumber = "ACTION" & act.ActionNumber
        .MetadataNumber = act.MetadataNumber
        .FromID = meta.FromID
        .ToID = meta.ToID
        .CacheRoot = RetLocalDirectory(locRoot & rmt.CacheRoot & meta.Directory)
        .BackupRoot = RetLocalDirectory(locRoot & rmt.BackupRoot & meta.Directory)
        .Root = meta.Directory
    End With
    
    GetFilename data, meta
    DoActionByMetadata act.dllName, data
End Sub

Private Function GetFilename(data As TYPE_ACTION_DATA, meta As TYPE_METADATA) As String
    Dim sFilename As String, subName As String
    Dim bRet As Boolean, Index As Long
    
    If Len(meta.DtFormat) > 0 Then
        sFilename = meta.Filename & Format(Now, meta.DtFormat)
    End If
    
    subName = TB_Runtime.RetUserDefine("TASK SN", meta.ID, "0")
    If Not IsNumeric(subName) Or Len(subName) > Len(meta.SnFormat) Then
        subName = "0"
    End If
    Index = CLng(subName): subName = ""
    
    Do
        
        Index = Index + 1
        If Len(CStr(Index)) > Len(meta.SnFormat) Then Index = 0
        
        subName = Format(Index, String(Len(meta.SnFormat), "0")) & "." & meta.Postfix
        If Len(Dir(data.CacheRoot & sFilename & subName, vbArchive)) = 0 Then
            bRet = True
        End If

    Loop While Not bRet
    
    TB_Runtime.SetUserDefine "TASK SN", meta.ID, Index
    data.Filename = sFilename & subName
End Function

Private Sub DoActionByMetadata(dllName As String, data As TYPE_ACTION_DATA)
    Dim task As TB_Runtime.ITask
    Dim sha As TYPE_XZ_SHARE
On Error GoTo HERROR

    If Len(dllName) > 0 Then
    
        sha.Resource = data.CacheRoot & data.Filename
        
        If Not ThisShare.IsLock(sha) Then
            ThisShare.DoLock sha
            
            Set task = CreateObject(dllName)
            If task.Init(ThisK3Server, data, ThisParamter) Then
                If task.Run(ThisParamter) Then
                    If data.ToID = ENUM_FTPROOT.ROOT_K3 Then
                        TB_Runtime.MakeDirByLocal data.BackupRoot
                        If ThisFileSystem.FileExists(data.BackupRoot & data.Filename) Then
                            ThisFileSystem.MoveFile data.BackupRoot & data.Filename, data.BackupRoot & data.Filename & ".bak" & Format(Now, "YYYYMMDDhhmmss")
                        End If
                        ThisFileSystem.MoveFile data.CacheRoot & data.Filename, data.BackupRoot & data.Filename
                    End If
                End If
            End If
            
            task.Dispose ThisParamter
            Set task = Nothing
            
            ThisShare.DeLock sha
        End If
    End If
    Exit Sub
    
HERROR:
    If ThisShare.IsLock(sha) Then
        ThisShare.DeLock sha
    End If
    mFunc.LogEx Err.Description
End Sub

Private Function RetRemoteDirectory(rmtDir As String) As String
    Dim rmtDir2 As String
    rmtDir2 = Replace(rmtDir, "\", "/")
    rmtDir2 = IIf(Left(rmtDir2, 1) <> "/", "/" & rmtDir2, rmtDir2)
    RetRemoteDirectory = rmtDir2
End Function

Private Function RetLocalDirectory(locDir As String) As String
    RetLocalDirectory = Replace(locDir, "/", "\")
End Function

Private Function SameNameString(meta As TYPE_METADATA, Filename As String) As Boolean
    If meta.Filename = Left(Filename, Len(meta.Filename)) And meta.Postfix = Right(Filename, Len(meta.Postfix)) Then
        If Len(meta.SnFormat) > 0 Then
            If Len(meta.Filename) + Len(meta.DtFormat) + Len(meta.SnFormat) + Len(meta.Postfix) + 1 = Len(Filename) Then
                SameNameString = True
            End If
        Else
            SameNameString = True
        End If
    End If
End Function
