VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ThisSize As Long
Private ThisAction() As TYPE_ACTION

Private Sub Class_Initialize()
    ThisSize = 0
End Sub

Private Sub Class_Terminate()
    Erase ThisAction
End Sub

Public Property Get Action(Index As Long) As TYPE_ACTION
    Action = ThisAction(Index)
End Property
Public Property Let Action(Index As Long, act As TYPE_ACTION)
    ThisAction(Index) = act
End Property

Public Property Get Size() As Long
    Size = ThisSize
End Property

Public Function Add(act As TYPE_ACTION) As Boolean
    Dim i As Long

    ReDim Preserve ThisAction(ThisSize)
    ThisAction(ThisSize) = act
    ThisSize = ThisSize + 1
    
    Add = True
End Function

Public Function Remove(Index As Long) As Boolean
    Dim i As Long
    
    If Index >= 0 And Index < ThisSize Then
        
        For i = Index To ThisSize - 2
            ThisAction(i) = ThisAction(i + 1)
        Next i
        ReDim Preserve ThisAction(ThisSize - 2)
        ThisSize = ThisSize - 1
        Remove = True
    Else
        Remove = False
    End If
End Function

