VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBMetadatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ThisSize As Long
Private ThisMetadata() As TYPE_METADATA

Private Sub Class_Initialize()
    ThisSize = 0
End Sub

Private Sub Class_Terminate()
    Erase ThisMetadata
End Sub

Public Property Get Metadata(Index As Long) As TYPE_METADATA
    Metadata = ThisMetadata(Index)
End Property
Public Property Let Metadata(Index As Long, meta As TYPE_METADATA)
    ThisMetadata(Index) = meta
End Property

Public Property Get Size() As Long
    Size = ThisSize
End Property

Public Function Lookup(ID As String) As Long
    Dim i As Long
    For i = 1 To ThisSize
        If ThisMetadata(i - 1).ID = ID Then
            Lookup = i - 1
            Exit Function
        End If
    Next i
    
    Lookup = mParam.CONST_ERROR_INDEX
End Function

Public Function Add(meta As TYPE_METADATA) As Boolean
    Dim i As Long
    i = Lookup(meta.ID)
    If i > -1 Then
        ThisMetadata(i) = meta
    Else
        ReDim Preserve ThisMetadata(ThisSize)
        ThisMetadata(ThisSize) = meta
        ThisSize = ThisSize + 1
    End If
    
    Add = True
End Function

Public Function Remove(Index As Long) As Boolean
    Dim i As Long
    
    If Index >= 0 And Index < ThisSize Then
        
        For i = Index To ThisSize - 2
            ThisMetadata(i) = ThisMetadata(i + 1)
        Next i
        ReDim Preserve ThisMetadata(ThisSize - 2)
        ThisSize = ThisSize - 1
        Remove = True
    Else
        Remove = False
    End If
End Function
