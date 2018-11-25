Attribute VB_Name = "mFunc"
Option Explicit

'Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As String, ByVal Source As Long, ByVal Length As Long)
'
'Public Function GetCommand() As String
'    Dim lRet As Long, sCmd As String
'    lRet = GetCommandLine
'    If lRet > 0 Then
'        sCmd = String(256, 32)
'        CopyMemory sCmd, lRet, Len(sCmd)
'        sCmd = Mid(sCmd, 1, InStr(1, sCmd, Chr(0)) - 1)
'    End If
'    GetCommand = sCmd
'End Function

Public Function GetRunParameter(cmdString As String) As TB_Context.TBParameters
    Dim params As TB_Context.TBParameters, param As TYPE_PARAMETER
    Dim l1 As Long, l2 As Long
    Dim tmpCmd As String, vCmd() As String

    Set params = New TB_Context.TBParameters

    If Len(cmdString) > 0 Then
        l1 = InStr(1, cmdString, "@")
        If l1 = 1 Then
            Do
                l2 = InStr(l1 + 1, cmdString, "@")
                If l2 = 0 Then
                    tmpCmd = Mid(cmdString, l1, Len(cmdString) + 1 - l1)
                Else
                    tmpCmd = Mid(cmdString, l1, l2 - l1)
                End If
                
                vCmd = Split(tmpCmd, ":")
                If UboundEx(vCmd) = 1 Then
                    param.Key = vCmd(0)
                    param.Value = vCmd(1)
                    params.Add param
                End If
                Erase vCmd
                
                l1 = l2
            Loop While l1 > 0
        End If
    End If
    
    Set GetRunParameter = params
    Set params = Nothing
End Function

Public Function GetRunCommand(params As TBParameters) As String
    Dim l1 As Long
    Dim cmdString As String
    Dim param As TYPE_PARAMETER

    For l1 = 0 To params.Size - 1
        param = params.Parameter(l1)
        cmdString = cmdString & param.Key & ":" & param.Value
    Next l1
    
    GetRunCommand = cmdString
End Function

Public Sub LogEx(ByVal Info As String)
    Dim TaskNumber As String
    TaskNumber = "TASK" & mStart.Parameter.Value("@TASK")
    TB_Runtime.Log TaskNumber, Info
End Sub
