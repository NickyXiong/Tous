VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XZMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const FILE_MAP_WRITE = &H2
Private Const FILE_MAP_READ = &H4
Private Const PAGE_READWRITE = 4&
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const CREATE_ALWAYS = 2
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (DesStr As Any, SrcStr As Any, ByVal MaxLen As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappingAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long

Private Const CONST_MAP_FILENAME As String = "Setting\filemap.page"
Private Const CONST_MAP_NAME As String = "XZ_FUN4_FILEMAP"
Private Const CONST_MAP_BUFFERINDEX  As Long = 16
Private Const CONST_MAP_BUFFERLENGTH As Long = 1024

Private ThisFileHandle As Long
Private ThisMapHandle As Long
Private ThisMapAddress As Long

Public Property Get Size() As Long
    Size = CONST_MAP_BUFFERLENGTH
End Property

Public Property Get Count() As Long
    Count = CONST_MAP_BUFFERINDEX
End Property

Public Function NewMemory() As Boolean
    Dim sBuffer As String
    Dim sMapFile As String
    
    sMapFile = IIf(Right(App.path, 1) = "\", App.path & CONST_MAP_FILENAME, App.path & "\" & CONST_MAP_FILENAME)
    
    ThisFileHandle = CreateFile(sMapFile, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If ThisFileHandle > 0 Then
        sBuffer = String(CONST_MAP_BUFFERINDEX * CONST_MAP_BUFFERLENGTH, " ")
        WriteFile ThisFileHandle, ByVal sBuffer, CONST_MAP_BUFFERINDEX * CONST_MAP_BUFFERLENGTH, 0, 0
        FlushFileBuffers ThisFileHandle
        
        ThisMapHandle = CreateFileMapping(ThisFileHandle, 0, PAGE_READWRITE, 0, 0, CONST_MAP_NAME)
        ThisMapAddress = MapViewOfFile(ThisMapHandle, FILE_MAP_WRITE, 0, 0, 0)
        
        If Not ThisMapAddress > 0 Then
            CloseHandle ThisFileHandle
            Exit Function
        End If
        NewMemory = True
    Else
        NewMemory = OpenMemory
    End If
End Function

Public Function OpenMemory() As Boolean
    ThisFileHandle = OpenFileMapping(FILE_MAP_WRITE, False, CONST_MAP_NAME)
    If Not ThisFileHandle > 0 Then Exit Function
    ThisMapAddress = MapViewOfFile(ThisFileHandle, FILE_MAP_WRITE, 0, 0, 0)
    
    If Not ThisMapAddress > 0 Then
        CloseHandle ThisFileHandle
        Exit Function
    End If
    
    OpenMemory = True
End Function

Public Sub Dispose()
    If ThisMapHandle > 0 Then
        CloseHandle ThisMapHandle
    End If
    If ThisFileHandle > 0 Then
        CloseHandle ThisFileHandle
    End If
End Sub

Public Function InputMemory(Index As Long) As String
    Dim sMessage As String
    
    If Index <= CONST_MAP_BUFFERINDEX And Index > 0 And ThisMapAddress > 0 Then
        sMessage = String(CONST_MAP_BUFFERLENGTH, " ")
        lstrcpyn ByVal sMessage, ByVal ThisMapAddress + (Index - 1) * CONST_MAP_BUFFERLENGTH, CONST_MAP_BUFFERLENGTH
        sMessage = Trim(sMessage)
        sMessage = IIf(Len(sMessage) > 0, Left(sMessage, Len(sMessage) - 1), "")
        InputMemory = Trim(sMessage)
    End If
End Function

Public Function OutputMemory(Index As Long, sMessage As String)
    Dim s2 As String
    If Index <= CONST_MAP_BUFFERINDEX And Index > 0 And ThisMapAddress > 0 Then
        s2 = sMessage & String(CONST_MAP_BUFFERLENGTH - Len(sMessage), " ")
        lstrcpyn ByVal ThisMapAddress + (Index - 1) * CONST_MAP_BUFFERLENGTH, ByVal s2, CONST_MAP_BUFFERLENGTH
    End If
End Function

