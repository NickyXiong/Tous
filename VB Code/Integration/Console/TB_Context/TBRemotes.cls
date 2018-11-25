VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBRemotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ThisSize As Long
Private ThisRemote() As TYPE_REMOTE

Private Sub Class_Initialize()
    ThisSize = 0
End Sub

Private Sub Class_Terminate()
    Erase ThisRemote
End Sub

Public Property Get Remote(Index As Long) As TYPE_REMOTE
    Remote = ThisRemote(Index)
End Property
Public Property Let Remote(Index As Long, rmt As TYPE_REMOTE)
    ThisRemote(Index) = rmt
End Property

Public Property Get Size() As Long
    Size = ThisSize
End Property

Public Function Lookup(RootID As ENUM_FTPROOT) As Long
    Dim i As Long
    For i = 1 To ThisSize
        If ThisRemote(i - 1).RootID = RootID Then
            Lookup = i - 1
            Exit Function
        End If
    Next i
    
    Lookup = mParam.CONST_ERROR_INDEX
End Function

Public Function Add(rmt As TYPE_REMOTE) As Boolean
    Dim i As Long
    i = Lookup(rmt.RootID)
    If i > -1 Then
        ThisRemote(i) = rmt
    Else
        ReDim Preserve ThisRemote(ThisSize)
        ThisRemote(ThisSize) = rmt
        ThisSize = ThisSize + 1
    End If
    
    Add = True
End Function

Public Function Remove(Index As Long) As Boolean
    Dim i As Long
    
    If Index >= 0 And Index < ThisSize Then
        
        For i = Index To ThisSize - 2
            ThisRemote(i) = ThisRemote(i + 1)
        Next i
        ReDim Preserve ThisRemote(ThisSize - 2)
        ThisSize = ThisSize - 1
        Remove = True
    Else
        Remove = False
    End If
End Function
