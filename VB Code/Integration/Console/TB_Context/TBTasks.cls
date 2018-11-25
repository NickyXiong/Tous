VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ThisSize As Long
Private ThisTask() As TYPE_TASK

Private Sub Class_Initialize()
    ThisSize = 0
End Sub

Private Sub Class_Terminate()
    Erase ThisTask
End Sub

Public Property Get Task(Index As Long) As TYPE_TASK
    Task = ThisTask(Index)
End Property
Public Property Let Task(Index As Long, Task As TYPE_TASK)
    ThisTask(Index) = Task
End Property

Public Property Get Size() As Long
    Size = ThisSize
End Property

Public Function Lookup(Number As String) As Long
    Dim i As Long
    For i = 1 To ThisSize
        If ThisTask(i - 1).Number = Number Then
            Lookup = i - 1
            Exit Function
        End If
    Next i
    
    Lookup = mParam.CONST_ERROR_INDEX
End Function

Public Function Add(Task As TYPE_TASK) As Boolean
    Dim i As Long
    i = Lookup(Task.Number)
    If i > -1 Then
        ThisTask(i) = Task
    Else
        ReDim Preserve ThisTask(ThisSize)
        ThisTask(ThisSize) = Task
        ThisSize = ThisSize + 1
    End If
    
    Add = True
End Function

Public Function Remove(Index As Long) As Boolean
    Dim i As Long
    
    If Index >= 0 And Index < ThisSize Then
        
        For i = Index To ThisSize - 2
            ThisTask(i) = ThisTask(i + 1)
        Next i
        ReDim Preserve ThisTask(ThisSize - 2)
        ThisSize = ThisSize - 1
        Remove = True
    Else
        Remove = False
    End If
End Function
