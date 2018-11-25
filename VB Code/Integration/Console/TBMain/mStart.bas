Attribute VB_Name = "mStart"
Option Explicit

Public Sub Main()
    If App.PrevInstance Then
        MsgBox "The program is running£¡", vbInformation + vbOKOnly, mParam.CONST_RUN_TITLE
        End
    End If

    Load frmMain
    Load frmRuntime
End Sub

Public Sub Dispose()
    Unload frmRuntime
    Unload frmMain
    End
End Sub
