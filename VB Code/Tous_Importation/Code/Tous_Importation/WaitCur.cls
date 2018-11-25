VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaitCur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private iOnCnt As Integer
Private iOldMousePtr As Integer

Private Sub Class_Initialize()
    SetHourglass True
End Sub

Private Sub Class_Terminate()
    SetHourglass False
End Sub

'*******************************************************************************
'* SUB: SetHourglass
'*    Provides a safer way to turn on and off the hourglass mouse pointer.
'*
'* PARAMETERS:
'*    fOn   If true the mouse cursor will be set to the hourglass cursor,
'*          if false the previous mouse cursor will be restored
'*******************************************************************************
Private Sub SetHourglass(fOn As Boolean)
      
    If fOn Then
        iOnCnt = iOnCnt + 1
        If iOnCnt = 1 Then
            iOldMousePtr = Screen.MousePointer
            Screen.MousePointer = vbHourglass
        End If
    Else
        iOnCnt = iOnCnt - 1
        If iOnCnt < 1 Then
            iOnCnt = 0
            Screen.MousePointer = iOldMousePtr
        End If
    End If
      
End Sub

