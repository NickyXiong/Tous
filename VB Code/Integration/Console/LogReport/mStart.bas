Attribute VB_Name = "mStart"
Option Explicit

Public Parameter As TB_Context.TBParameters
Public Is_Exit As Boolean
Public Is_CommandExit As Boolean

Public Sub Main()
    Dim bRun As Boolean
    Dim sCommand As String
On Error GoTo HERROR

    If App.EXEName = "LogReport" Then
        
        sCommand = Command
        Set Parameter = mFunc.GetRunParameter(sCommand)

        If Parameter.Size > 0 Then
            If Parameter.Lookup("@TASK") > -1 And Parameter.Lookup("@ACTION") > -1 And Parameter.Lookup("@MASK") > -1 Then
                bRun = True
            End If
        End If
    End If
    
    If bRun Then
        mFunc.LogEx "LogReport Running..."
        DoAction
        mFunc.LogEx "LogReport End"
    Else
        CopyRight.Show vbModal
    End If
    End
HERROR:
    mFunc.LogEx Err.Description
    mFunc.LogEx "FileSystem End"
    End
End Sub

Private Sub DoAction()
    Dim ilod As TB_Runtime.ILoad
    Dim param As TB_Context.TBParameters
    
    Load CopyRight
    CopyRight.Caption = Parameter.Value("@MASK")
    
    Set ilod = New TBSend
    Set param = New TB_Context.TBParameters
    If ilod.Init(param) Then
        ilod.Run param
    End If
    ilod.Dispose param
    Set ilod = Nothing
    Set param = Nothing
    
    Is_Exit = True
    Unload CopyRight
End Sub
