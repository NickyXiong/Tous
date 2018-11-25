VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XZProcessEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const WM_CLOSE = &H10
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
 End Type
 
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function CoCreateGuid Lib "ole32.dll" (tGUIDStructure As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Private ThisProcess() As TYPE_XZ_PROCESSATTRIBUTE
Private ThisProcessCount As Long

Public Property Get Process(Index As Long) As TYPE_XZ_PROCESSATTRIBUTE
    Process = ThisProcess(Index)
End Property

Public Property Get Status(Index As Long) As ENUM_XZ_PROCESSSTATE
    Status = GetStatusByIndex(Index)
End Property

Public Function NewProcess(sKey As String, sDir As String, Optional sCommand As String = "") As Long
    Dim lIndex As Long
    lIndex = Lookup(sKey)
    If lIndex = mParam.CONST_ERROR_INDEX Then
        lIndex = ThisProcessCount
        ReDim Preserve ThisProcess(lIndex)
        ThisProcess(lIndex).Key = sKey
        ThisProcess(lIndex).Directory = sDir
        ThisProcess(lIndex).Command = sCommand
        ThisProcess(lIndex).Mask = GetMask
        ThisProcessCount = lIndex + 1
        NewProcess = lIndex
        Exit Function
    End If
    NewProcess = mParam.CONST_ERROR_INDEX
End Function

Public Function RunProcess(sKey As String, sApp As String, Optional sCommand As String = "") As Long
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
    ThisProcess(lIndex).Mask = GetMask
    
    If GetStatusByIndex(lIndex) = ENUM_XZ_PROCESSSTATE.STATE_RUN Then
        StopProcessByIndex lIndex
    End If
    StartProcessByIndex lIndex
    RunProcess = lIndex
End Function

Private Function GetMask() As String
    Dim sGuid As String
    Dim tGuid As GUID
    Dim bGuid() As Byte
    Dim lRtn As Long
    Const clLen As Long = 50
   
    If CoCreateGuid(tGuid) = 0 Then
       bGuid = String(clLen, 0)
       lRtn = StringFromGUID2(tGuid, VarPtr(bGuid(0)), clLen)
       If lRtn > 0 Then
          sGuid = Mid$(bGuid, 1, lRtn - 1)
       End If
    End If
    
    If Len(sGuid) = 0 Then
        sGuid = Format(Now, "YYYYMMDDhhmmss")
    End If
    GetMask = sGuid
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

Public Sub StopProcess(sKey As String)
    Dim lIndex As Long
    lIndex = Lookup(sKey)
    If lIndex >= mParam.CONST_ERROR_INDEX Then
        StopProcessByIndex lIndex
    End If
End Sub

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
    ThisProcess(Index).PHD = FindWindow(vbNullString, ThisProcess(Index).Mask)
    If ThisProcess(Index).PHD > 0 Then
        GetStatusByIndex = ENUM_XZ_PROCESSSTATE.STATE_RUN
        Exit Function
    End If
    GetStatusByIndex = ENUM_XZ_PROCESSSTATE.STATE_STOP
End Function

Private Function StartProcessByIndex(Index As Long) As Boolean
    ThisProcess(Index).PHD = Shell(ThisProcess(Index).Directory & " " & ThisProcess(Index).Command & "@MASK:" & ThisProcess(Index).Mask)
    If ThisProcess(Index).PHD > 0 Then
        StartProcessByIndex = True
    End If
End Function

Private Function StopProcessByIndex(Index As Long) As Boolean
    ThisProcess(Index).PHD = FindWindow(vbNullString, ThisProcess(Index).Mask)
    If ThisProcess(Index).PHD > 0 Then
        SendMessage ThisProcess(Index).PHD, WM_CLOSE, 0, 0
    End If
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
