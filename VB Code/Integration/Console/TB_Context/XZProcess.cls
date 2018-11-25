VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XZProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Private ThisProcess() As TYPE_XZ_PROCESSATTRIBUTE
Private ThisProcessCount As Long

Public Property Get Process(Index As Long) As TYPE_XZ_PROCESSATTRIBUTE
    Process = ThisProcess(Index)
End Property

Public Property Get Status(Index As Long) As ENUM_XZ_PROCESSSTATE
    Status = GetStatusByIndex(Index)
End Property

Public Function GetCurrentID() As Long
    GetCurrentID = GetCurrentProcessId
End Function

Public Function NewProcess(sKey As String, sDir As String, Optional sCommand As String = "") As Long
    Dim lIndex As Long
    lIndex = Lookup(sKey)
    If lIndex = mParam.CONST_ERROR_INDEX Then
        lIndex = ThisProcessCount
        ReDim Preserve ThisProcess(lIndex)
        ThisProcess(lIndex).Key = sKey
        ThisProcess(lIndex).Directory = sDir
        ThisProcess(lIndex).Command = sCommand
        ThisProcessCount = lIndex + 1
        NewProcess = lIndex
        Exit Function
    End If
    NewProcess = mParam.CONST_ERROR_INDEX
End Function

Public Function RunProcess(sKey As String, sApp As String, Optional sCommand As String = "") As Long
    Dim ppAttr As SECURITY_ATTRIBUTES
    Dim ptAttr As SECURITY_ATTRIBUTES
    Dim psInfo As STARTUPINFO
    Dim ppInfo As PROCESS_INFORMATION
    Dim lpt As Long, lIndex As Long
    
    lIndex = Lookup(sKey)
    If lIndex = mParam.CONST_ERROR_INDEX Then
        ReDim Preserve ThisProcess(ThisProcessCount)
        lIndex = ThisProcessCount
        ThisProcessCount = ThisProcessCount + 1
    End If
    
    ThisProcess(lIndex).Key = sKey
    ThisProcess(lIndex).Directory = sApp
    ThisProcess(lIndex).Command = sCommand
    
    If GetStatusByIndex(lIndex) = ENUM_XZ_PROCESSSTATE.STATE_RUN Then
        StopProcessByIndex lIndex
    End If
    StartProcessByIndex lIndex
    RunProcess = lIndex
End Function

Public Function StartProcess(sKey As String) As Boolean
    Dim lIndex As Long
    lIndex = Lookup(sKey)
    If lIndex >= mParam.CONST_ERROR_INDEX Then
        If GetStatusByIndex(lIndex) <> ENUM_XZ_PROCESSSTATE.STATE_RUN Then
            StartProcess = StartProcessByIndex(lIndex)
        Else
            StartProcess = True
        End If
    End If
End Function

Public Function StopProcess(sKey As String) As Boolean
    Dim lIndex As Long
    lIndex = Lookup(sKey)
    If lIndex >= mParam.CONST_ERROR_INDEX Then
        StopProcess = StopProcessByIndex(lIndex)
    End If
End Function

Public Function RemoveProcess(sKey As String) As Boolean
    Dim lIndex As Long
    lIndex = Lookup(sKey)
    If lIndex >= mParam.CONST_ERROR_INDEX Then
        StopProcessByIndex lIndex
        RemoveProcess = RemoveProcessByIndex(lIndex)
    End If
End Function

Public Function GetStatus(sKey As String) As ENUM_XZ_PROCESSSTATE
    Dim lIndex As Long
    lIndex = Lookup(sKey)
    If lIndex >= mParam.CONST_ERROR_INDEX Then
        GetStatus = GetStatusByIndex(lIndex)
    Else
        GetStatus = ENUM_XZ_PROCESSSTATE.STATE_NOMAL
    End If
End Function

Private Function GetStatusByIndex(Index As Long) As ENUM_XZ_PROCESSSTATE
    Dim lpt As Long, lpid As Long
    
    If ThisProcess(Index).PHD > 0 Then
        lpt = OpenProcess(PROCESS_ALL_ACCESS, False, ThisProcess(Index).PID)
        If lpt Then
            GetStatusByIndex = ENUM_XZ_PROCESSSTATE.STATE_RUN
            Exit Function
        End If
    End If
    GetStatusByIndex = ENUM_XZ_PROCESSSTATE.STATE_STOP
End Function

Private Function StartProcessByIndex(Index As Long) As Boolean
    Dim ppAttr As SECURITY_ATTRIBUTES
    Dim ptAttr As SECURITY_ATTRIBUTES
    Dim psInfo As STARTUPINFO
    Dim ppInfo As PROCESS_INFORMATION
    Dim lpt As Long

    lpt = CreateProcess(ThisProcess(Index).Directory, ThisProcess(Index).Command, ppAttr, ptAttr, 0, 0, 0, App.path, psInfo, ppInfo)
    If lpt Then
        ThisProcess(Index).PHD = ppInfo.hProcess
        ThisProcess(Index).PID = ppInfo.dwProcessId
        ThisProcess(Index).THD = ppInfo.hThread
        ThisProcess(Index).TID = ppInfo.dwThreadId
        StartProcessByIndex = True
    End If
End Function

Private Function StopProcessByIndex(Index As Long) As Boolean
    If ThisProcess(Index).PHD > 0 Then
        TerminateProcess ThisProcess(Index).PHD, 0&
    End If
    If ThisProcess(Index).THD > 0 Then
        CloseHandle ThisProcess(Index).THD
    End If
    If ThisProcess(Index).PHD > 0 Then
        CloseHandle ThisProcess(Index).THD
    End If
    ThisProcess(Index).PHD = 0
    ThisProcess(Index).PID = 0
    ThisProcess(Index).THD = 0
    ThisProcess(Index).TID = 0
    
    StopProcessByIndex = True
End Function

Private Function RemoveProcessByIndex(Index As Long) As Boolean
    Dim lIndex As Long
    For lIndex = Index To ThisProcessCount - 2
        ThisProcess(lIndex) = ThisProcess(lIndex + 1)
    Next lIndex
    ThisProcessCount = ThisProcessCount - 1
    If ThisProcessCount = 0 Then
        Erase ThisProcess
    Else
        ReDim Preserve ThisProcess(ThisProcessCount - 1)
    End If
    
    RemoveProcessByIndex = True
End Function

Public Function GetProcess(sKey As String) As TYPE_XZ_PROCESSATTRIBUTE
    Dim lIndex As Long
    lIndex = Lookup(sKey)
    If lIndex >= mParam.CONST_ERROR_INDEX Then
        GetProcess = ThisProcess(lIndex)
    End If
End Function

Public Function Lookup(sKey As String) As Long
    Dim lIndex As Long
    For lIndex = 0 To ThisProcessCount - 1
        If ThisProcess(lIndex).Key = sKey Then
            Lookup = lIndex
            Exit Function
        End If
    Next
    Lookup = mParam.CONST_ERROR_INDEX
End Function

Public Sub Dispose()
    Dim lIndex As Long
    For lIndex = 0 To ThisProcessCount - 1
        StopProcessByIndex lIndex
    Next lIndex
End Sub

Private Sub Class_Terminate()
    Dispose
    Erase ThisProcess
End Sub
