VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements TB_Runtime.ILoad

Private ThisRemotes As TB_Context.TBRemotes
Private ThisInternet As XZInternet
Private ThisShare As TB_Context.XZShareLock
Private ThisMetadatas As TB_Context.TBMetadatas
Private ThisFileSystem As FileSystemObject

Private SFTP_strServer As String
Private SFTP_strUsername As String
Private SFTP_strPassword As String
Private SFTP_strDownload As String
Private SFTP_strUpload As String
Private SFTP_strPort As String

Private Function ILoad_Init(param As TB_Context.TBParameters) As Boolean
    Dim strReturn As String
    
On Error GoTo HERROR

    mFunc.LogEx "File Download Initialize"
    DoEvents
    
    Dim oLoc As TB_Runtime.LocalContent
    Set oLoc = New TB_Runtime.LocalContent
    Set ThisRemotes = oLoc.GetRemotes
    Set ThisMetadatas = oLoc.GetMetadatas
    Set oLoc = Nothing

    Set ThisInternet = New XZInternet
    Set ThisShare = New TB_Context.XZShareLock
    Set ThisFileSystem = New FileSystemObject
    
    
    '读取SFTP配置
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "Server", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    SFTP_strServer = Replace(strReturn, Chr(0), "")
    
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "Username", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    SFTP_strUsername = Replace(strReturn, Chr(0), "")
    
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "Password", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    SFTP_strPassword = Replace(strReturn, Chr(0), "")
    
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "DownloadRoot", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    SFTP_strDownload = Replace(strReturn, Chr(0), "")
    
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "UploadRoot", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    SFTP_strUpload = Replace(strReturn, Chr(0), "")
    
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "Port", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    SFTP_strPort = Replace(strReturn, Chr(0), "")
    
    ILoad_Init = True
    Exit Function
HERROR:
    mFunc.LogEx Err.Description
End Function

Private Sub ILoad_Run(param As TB_Context.TBParameters)
    mFunc.LogEx "File Download Running"
    DoEvents
    
    DownloadByRemote
End Sub

Private Sub ILoad_Dispose(param As TB_Context.TBParameters)
    mFunc.LogEx "File Download Dipose"
    Set ThisInternet = Nothing
    Set ThisShare = Nothing
    Set ThisRemotes = Nothing
    Set ThisMetadatas = Nothing
    Set ThisFileSystem = Nothing
End Sub

Private Sub DownloadByRemote()
    Dim lIndex As Long, bRet As Boolean
    Dim rmt As TYPE_REMOTE, meta As TYPE_METADATA
On Error GoTo HERROR

    lIndex = ThisRemotes.Lookup(ENUM_FTPROOT.ROOT_3PL Or ENUM_FTPROOT.ROOT_POS Or ENUM_FTPROOT.ROOT_TBHQ)
    If lIndex = -1 Then Err.Raise -1, "", "The remote server is not set"
    
    rmt = ThisRemotes.Remote(lIndex)
    mFunc.LogEx "connection remote server " & rmt.Server
    bRet = ThisInternet.Connection(rmt.Server, rmt.Username, rmt.Password)
    If Not bRet Then Err.Raise -1, "", "On the remote server failure!"
    
    For lIndex = 0 To ThisMetadatas.Size - 1
        If mStart.Is_CommandExit Then
            Exit For
        End If
        
        meta = ThisMetadatas.Metadata(lIndex)
'        If Not mStart.Is_CommandExit And meta.FromID <> ENUM_FTPROOT.ROOT_K3 Then
'            DownloadByMetadatas rmt, meta
'        End If
        If Not mStart.Is_CommandExit And meta.ToID = ENUM_FTPROOT.ROOT_K3 Then
            If meta.FromID = ENUM_FTPROOT.ROOT_TBHQ Then
                DownloadByMetadatas rmt, meta
            ElseIf meta.FromID = ENUM_FTPROOT.ROOT_POS Then
                mFunc.LogEx "connection SFTP server " & rmt.Server
                SFTPDownloadByMetadatas rmt, meta
            End If
        End If
        
        DoEvents
    Next lIndex
    
    ThisInternet.Dispose
    Exit Sub
HERROR:
    mFunc.LogEx Err.Description
    ThisInternet.Dispose
End Sub

Private Sub DownloadByMetadatas(rmt As TYPE_REMOTE, meta As TYPE_METADATA)
    Dim lIndex As Long
    Dim vFiles() As String
    Dim rmtDir As String, cacheDir As String
    Dim locRoot As String
    
    locRoot = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    rmtDir = RetRemoteDirectory(rmt.DownRoot & meta.Directory)
    vFiles = ThisInternet.GetFiles(rmtDir)
    
    For lIndex = 0 To TB_Runtime.UboundEx(vFiles)
        If mStart.Is_CommandExit Then
            Exit For
        End If
    
        If SameNameString(meta, vFiles(lIndex)) Then
        
            rmtDir = RetRemoteDirectory(rmt.DownRoot & meta.Directory & vFiles(lIndex))
            cacheDir = RetLocalDirectory(locRoot & rmt.CacheRoot & meta.Directory & vFiles(lIndex))
            
            DownloadByMetadata rmtDir, cacheDir
        End If
        
        DoEvents
    Next lIndex
End Sub

Private Function DownloadByMetadata(rmtDir As String, cacheDir As String) As Boolean
    Dim bRet As Boolean
    Dim sha As TYPE_XZ_SHARE
On Error GoTo HERROR

    sha.Resource = cacheDir
    
    If Not ThisShare.IsLock(sha) Then
        ThisShare.DoLock sha
        
        bRet = ThisInternet.Download(rmtDir, cacheDir)

        If bRet Then
            bRet = ThisFileSystem.FileExists(cacheDir)
        Else
            Err.Raise -1, "", "The file '" & rmtDir & "' download failure! The local file does not exist."
        End If

        If bRet Then
            If mStart.Parameter.Value("@DEL") = 1 Then
                'Tous HQ特殊处理，下载完成后将文件从pending目录移动至processed目录
                ThisInternet.Upload Replace(rmtDir, "pending", "processed"), cacheDir
                bRet = ThisInternet.RemoveFile(rmtDir)
            End If
        Else
            Err.Raise -1, "", "The file '" & rmtDir & "' download failure! Unknown causes after downloading the file is missing."
        End If

        mFunc.LogEx "The file '" & rmtDir & "' download success!"
        ThisShare.DeLock sha
    End If
    Exit Function
HERROR:
    If ThisShare.IsLock(sha) Then
        ThisShare.DeLock sha
    End If
    If Err.Number = -1 Then
        mFunc.LogEx Err.Description
    Else
        mFunc.LogEx "The file '" & rmtDir & "' download failure! " & Err.Description
    End If
End Function


Private Sub SFTPDownloadByMetadatas(rmt As TYPE_REMOTE, meta As TYPE_METADATA)
    Dim strCommand As String, strConfig As String
    Dim filenumber As Integer
    Dim lIndex As Long
    Dim vFiles() As String
    Dim rmtDir As String, cacheDir As String
    Dim locRoot As String
    Dim objFileList As Object, f As File
    Dim i As Integer
On Error GoTo HERROR
    
    locRoot = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    rmtDir = RetRemoteDirectory(rmt.DownRoot & meta.Directory)
    cacheDir = RetLocalDirectory(locRoot & rmt.CacheRoot & meta.Directory)
    MakeDirByLocal cacheDir
    
    '构建下载指令后通过psftp调用
    strCommand = "cd " & rmtDir
    strCommand = strCommand & vbCrLf & "lcd """ & cacheDir & """"
    strCommand = strCommand & vbCrLf & "mget " & meta.Filename & "*." & meta.Postfix
    strCommand = strCommand & vbCrLf & "rm " & meta.Filename & "*." & meta.Postfix
'    strCommand = strCommand & vbCrLf & "cd backup"
'    strCommand = strCommand & vbCrLf & "mput *." & strSuffix
    strCommand = strCommand & vbCrLf & "bye"

    '指令导出为txt文件，通过psftp执行
    filenumber = FreeFile
    strConfig = IIf(Right(App.path, 1) = "\", App.path, App.path & "\") & "Setting\SFTPDownloadCmd.txt"
    Open strConfig For Output As #filenumber
    Print #filenumber, strCommand
    Close #filenumber
    
    strCommand = """" & IIf(Right(App.path, 1) = "\", App.path, App.path & "\") & "psftp.exe " _
                & """" & " " & SFTP_strUsername & "@" & SFTP_strServer & " -P " & SFTP_strPort & " -pw " & SFTP_strPassword & " -b " & """" & strConfig & """"

    CloseHandle dcWaitForSingleObject(OpenProcess(SYNCHRONIZE, 0, Shell(strCommand, 0)), 15000)  '执行cmd.exe，并等待其结束。

    DoEvents
    
    '检查文件创建时间，1分钟内的文件名将被写入日志
    Set objFileList = ThisFileSystem.GetFolder(cacheDir)
    
    For Each f In objFileList.Files
        If DateDiff("n", CDate(f.DateCreated), Now()) < 3 And InStr(f.Name, meta.Filename) > 0 Then
            mFunc.LogEx "The file '" & f.Name & "' download success!"
        End If
    Next

    Exit Sub
HERROR:
    If Err.Number = -1 Then
        mFunc.LogEx Err.Description
    Else
        mFunc.LogEx "The file '" & rmtDir & "' download failure! " & Err.Description
    End If
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

Private Function MakeDirByLocal(ByVal path As String) As Boolean
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
