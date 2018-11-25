VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBMails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ThisSize As Long
Private ThisMails() As String

Private Sub Class_Initialize()
    ThisSize = 0
End Sub

Private Sub Class_Terminate()
    Erase ThisMails
End Sub

Public Property Get Mail(Index As Long) As String
    Mail = ThisMails(Index)
End Property
Public Property Let Mail(Index As Long, e As String)
    ThisMails(Index) = e
End Property

Public Property Get Size() As Long
    Size = ThisSize
End Property

Public Function Lookup(e As String) As Long
    Dim i As Long
    For i = 1 To ThisSize
        If ThisMails(i - 1) = e Then
            Lookup = i - 1
            Exit Function
        End If
    Next i
    
    Lookup = mParam.CONST_ERROR_INDEX
End Function

Public Function Add(e As String) As Boolean
    Dim i As Long
    i = Lookup(e)
    If i > -1 Then
        ThisMails(i) = e
    Else
        ReDim Preserve ThisMails(ThisSize)
        ThisMails(ThisSize) = e
        ThisSize = ThisSize + 1
    End If
    
    Add = True
End Function

Public Function Remove(Index As Long) As Boolean
    Dim i As Long
    
    If Index >= 0 And Index < ThisSize Then
        
        For i = Index To ThisSize - 2
            ThisMails(i) = ThisMails(i + 1)
        Next i
        ReDim Preserve ThisMails(ThisSize - 2)
        ThisSize = ThisSize - 1
        Remove = True
    Else
        Remove = False
    End If
End Function
