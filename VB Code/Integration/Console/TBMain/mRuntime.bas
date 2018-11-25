Attribute VB_Name = "mRuntime"
Option Explicit

Public Tasks As TB_Context.TBTasks
Public Process As TB_Context.XZProcessEx
Private ThisInterval() As Long
Private ThisIsRunend() As Long

Public Sub InitRuntime()
    Dim Index As Long, locDir As String
    Dim o As TB_Runtime.LocalContent
    Dim tsk As TYPE_TASK
    
    Set o = New TB_Runtime.LocalContent
    Set Tasks = o.GetTasks
    Set o = Nothing
    
    Set Process = New TB_Context.XZProcessEx
    locDir = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    ReDim ThisInterval(Tasks.Size - 1)
    ReDim ThisIsRunend(Tasks.Size - 1)
    
    With Process
        For Index = 0 To Tasks.Size - 1
            tsk = Tasks.task(Index)
            .NewProcess tsk.Number, locDir & tsk.EXEName, tsk.Parameter
        Next Index
    End With
End Sub

Public Function IsRun(ByVal Number As String) As Boolean
    Dim Index As Long
    
    Index = Process.Lookup(Number)
    If Index > -1 Then
        IsRun = IIf(Process.Status(Index) = ENUM_XZ_PROCESSSTATE.STATE_RUN, True, False)
    End If
End Function

Private Function IsRunByIndex(Index As Long) As Boolean
    IsRunByIndex = IIf(Process.Status(Index) = ENUM_XZ_PROCESSSTATE.STATE_RUN, True, False)
End Function

Public Function IsRunTime(ByVal Number As String, Optional ByVal Interval As Integer = 1) As Boolean
    Dim Index As Long
    
    Index = Tasks.Lookup(Number)
    If Index > -1 Then
        If Tasks.task(Index).RunStyle = ENUM_RUNSTYLE.RUNSTYLE_ACTUAL Then
            ThisInterval(Index) = ThisInterval(Index) + Interval
            If ThisInterval(Index) >= Tasks.task(Index).Interval Then
                IsRunTime = True
                ThisInterval(Index) = 0
            End If
        Else
            If InTimeRange(Tasks.task(Index).StartTime, DateAdd("n", 5, Tasks.task(Index).StartTime), Now) Then
                If ThisIsRunend(Index) = 0 Then
                    IsRunTime = True
                    ThisIsRunend(Index) = 1
                End If
            Else
                ThisIsRunend(Index) = 0
            End If
        End If
    End If
End Function

Public Sub RunTask(ByVal Number As String)
    Process.StartProcess Number
End Sub

Public Sub StopTask(ByVal Number As String)
    Process.StopProcess Number
End Sub

Private Function InTimeRange(ByVal BegTime As String, ByVal EndTime As String, ByVal ThisTime As String) As Boolean
    Dim lBeg As Long
    Dim lEnd As Long
    Dim lNow As Long
    
    lBeg = Val(Format(BegTime, "hhmm"))
    lEnd = Val(Format(EndTime, "hhmm"))
    lNow = Val(Format(ThisTime, "hhmm"))
    
    If (lNow + (2400 - lBeg)) Mod 2400 <= (lEnd + (2400 - lBeg)) Mod 2400 Then
        InTimeRange = True
    End If
End Function

Public Sub Dispose()
    Process.Dispose
    Set Tasks = Nothing
    Set Process = Nothing
End Sub
