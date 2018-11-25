VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ThisSize As Long
Private ThisParameter() As TYPE_PARAMETER

Private Sub Class_Initialize()
    ThisSize = 0
End Sub

Private Sub Class_Terminate()
    Erase ThisParameter
End Sub

Public Property Get Parameter(Index As Long) As TYPE_PARAMETER
    Parameter = ThisParameter(Index)
End Property
Public Property Let Parameter(Index As Long, param As TYPE_PARAMETER)
    ThisParameter(Index) = param
End Property

Public Property Get Key(sKey As String) As TYPE_PARAMETER
    Key = ThisParameter(Lookup(sKey))
End Property
Public Property Let Key(sKey As String, param As TYPE_PARAMETER)
    ThisParameter(Lookup(sKey)) = param
End Property

Public Property Get Value(sKey As String) As String
    Value = ThisParameter(Lookup(sKey)).Value
End Property
Public Property Let Value(sKey As String, sVal As String)
    ThisParameter(Lookup(sKey)).Value = sVal
End Property

Public Property Get Size() As Long
    Size = ThisSize
End Property

Public Function Lookup(sKey As String) As Long
    Dim i As Long
    For i = 1 To ThisSize
        If ThisParameter(i - 1).Key = sKey Then
            Lookup = i - 1
            Exit Function
        End If
    Next i
    
    Lookup = CONST_ERROR_INDEX
End Function

Public Function Add(param As TYPE_PARAMETER) As Boolean
    Dim i As Long
    i = Lookup(param.Key)
    If i > CONST_ERROR_INDEX Then
        ThisParameter(i) = param
    Else
        ReDim Preserve ThisParameter(ThisSize)
        ThisParameter(ThisSize) = param
        ThisSize = ThisSize + 1
    End If
    
    Add = True
End Function

Public Function Remove(Index As Long) As Boolean
    Dim i As Long
    
    If Index >= 0 And Index < ThisSize Then
        
        For i = Index To ThisSize - 2
            ThisParameter(i) = ThisParameter(i + 1)
        Next i
        ReDim Preserve ThisParameter(ThisSize - 2)
        ThisSize = ThisSize - 1
        Remove = True
    Else
        Remove = False
    End If
End Function

