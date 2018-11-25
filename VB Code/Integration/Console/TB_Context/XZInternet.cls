VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XZInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal LAccessType As Long, ByVal sProxyName As String, ByVal SProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExsiting As String, ByVal lpszNew As String) As Boolean

Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, lpszCurrentDirectory As String, ByRef lpdwCurrentDirectory As Long) As Boolean
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFndData As WIN32_FIND_DATA) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Private Const FILE_ATTRIBUTE_ENCRYPTED As Long = &H4000
Private Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_ATTRIBUTE_OFFLINE As Long = &H1000
Private Const FILE_ATTRIBUTE_READONLY As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Private Const INTERNET_OPEN_TYPE_PRECONFIG = &H0
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE As Long = &H4000000
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000
Private Const MAX_PATH = 260

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private ThisInternet As Long
Private ThisInternetConnection As Long

Private ThisHostname As String
Private ThisUsername As String
Private ThisPassword As String

Public Property Get State() As Boolean
    If ThisInternet > 0 And ThisInternetConnection > 0 Then
        State = True
    End If
End Property

Public Property Get Hostname() As String
    Hostname = ThisHostname
End Property
Public Property Get Username() As String
    Username = ThisUsername
End Property
Public Property Get Password() As String
    Password = ThisPassword
End Property

Public Function Connection(inetHostname As String, inetUsername As String, inetPassword As String) As Boolean
    Dispose
    ThisInternet = InternetOpen(vbNullString, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)
    If ThisInternet > 0 Then
        ThisInternetConnection = InternetConnect(ThisInternet, inetHostname, 0, inetUsername, inetPassword, 1, INTERNET_FLAG_PASSIVE, 0)
    End If
    If ThisInternetConnection > 0 Then
        Connection = True
        ThisHostname = inetHostname
        ThisUsername = inetUsername
        ThisPassword = inetPassword
    Else
        Dispose
    End If
End Function

Public Sub Dispose()
    If ThisInternetConnection > 0 Then
        InternetCloseHandle ThisInternetConnection
    End If
    If ThisInternet > 0 Then
        InternetCloseHandle ThisInternet
    End If
    
    ThisInternetConnection = 0
    ThisInternet = 0
    
    ThisHostname = ""
    ThisUsername = ""
    ThisPassword = ""
End Sub

Public Function Download(inetDir As String, localDir As String) As Boolean
    Dim bRet As Boolean
    
    If State Then
        MakeDirByLocal localDir
        bRet = FtpGetFile(ThisInternetConnection, inetDir, localDir, 0, 0, 1, 0)
    End If
    
    Download = bRet
End Function

Public Function Upload(inetDir As String, localDir As String) As Boolean
    Dim bRet As Boolean
    
    If State Then
        bRet = MakeDirectory(inetDir)
        If bRet Then
            bRet = FtpPutFile(ThisInternetConnection, localDir, inetDir, 1, 0)
        End If
    End If
    
    Upload = bRet
End Function

Public Function Rename(inetDir As String, inetDirNew As String) As Boolean
    If State Then
        Rename = FtpRenameFile(ThisInternetConnection, inetDir, inetDirNew)
    End If
End Function

Public Function RemoveFile(inetDir As String) As Boolean
    Dim bRet As Boolean

    If State Then
        bRet = FtpDeleteFile(ThisInternetConnection, inetDir)
    End If
    
    RemoveFile = bRet
End Function

Public Function MakeDirectory(inetDir As String) As Boolean
    Dim lIndex As Integer
    Dim bRet As Boolean
    Dim tempPath As String
On Error GoTo HERROR
    
    lIndex = InStr(1, inetDir, "/")
    While lIndex > 0
        tempPath = Left(inetDir, lIndex)

        FtpCreateDirectory ThisInternetConnection, tempPath
    
        lIndex = InStr(lIndex + 1, inetDir, "/")
    Wend

    MakeDirectory = True
    Exit Function
HERROR:
    MakeDirectory = False
End Function

Public Function GetFiles(inetDir As String) As String()
    Dim lpt As Long
    Dim bRet As Boolean
    Dim pData As WIN32_FIND_DATA
    Dim sTmp As String
    
    Dim lCount As Long
    Dim vFile() As String
    
    If State Then
        
        bRet = FtpSetCurrentDirectory(ThisInternetConnection, inetDir)
        lpt = FtpFindFirstFile(ThisInternetConnection, "*.*", pData, 0, 0)

        Do
            sTmp = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
            If Left(sTmp, 1) <> "." Then
                If pData.dwFileAttributes = FILE_ATTRIBUTE_NORMAL Then
                    ReDim Preserve vFile(lCount)
                    vFile(lCount) = sTmp
                    lCount = lCount + 1
                End If
            End If
            bRet = InternetFindNextFile(lpt, pData)

        Loop While bRet
        
        InternetCloseHandle lpt
    End If
    
    GetFiles = vFile
    Erase vFile
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

Private Sub Class_Terminate()
    Dispose
End Sub
