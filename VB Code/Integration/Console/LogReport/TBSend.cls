VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILoad

Private ThisActions As TB_Context.TBActions




Private Function ILoad_Init(param As TB_Context.TBParameters) As Boolean
On Error GoTo HERROR
    
    mFunc.LogEx "Task:" & mStart.Parameter.Value("@TASK") & " Action:" & mStart.Parameter.Value("@ACTION") & " Initialize"
    DoEvents
    
    Dim o As TB_Runtime.LocalContent
    Set o = New TB_Runtime.LocalContent
    Set ThisActions = o.GetActions
    Set o = Nothing
    
    ILoad_Init = True
    Exit Function
    
HERROR:
    mFunc.LogEx Err.Description
End Function

Private Sub ILoad_Run(param As TB_Context.TBParameters)
    mFunc.LogEx "Task:" & mStart.Parameter.Value("@TASK") & " Action:" & mStart.Parameter.Value("@ACTION") & " Run"
    DoEvents
    DoAction
End Sub

Private Sub ILoad_Dispose(param As TB_Context.TBParameters)
    mFunc.LogEx "Task:" & mStart.Parameter.Value("@TASK") & " Action:" & mStart.Parameter.Value("@ACTION") & " Dispose"
    DoEvents
    
    Set ThisActions = Nothing
End Sub

Private Sub DoAction()
    Dim Index As Long
    Dim act As TYPE_ACTION
On Error GoTo HERROR

    For Index = 0 To ThisActions.Size - 1
    
        If mStart.Is_CommandExit Then
            Exit For
        End If
    
        act = ThisActions.Action(Index)
        If mStart.Parameter.Value("@TASK") = act.TaskNumber And mStart.Parameter.Value("@ACTION") = act.ActionNumber Then
            DoActionEx act
        End If
        
        DoEvents
        
    Next Index
    Exit Sub
HERROR:
    mFunc.LogEx Err.Description
End Sub

Private Sub DoActionEx(act As TYPE_ACTION)
    Dim task As ILoad
On Error GoTo HERROR

    If Len(act.DllName) > 0 Then
        Set task = CreateObject(act.DllName)
        If task.Init(mStart.Parameter) Then
            task.Run mStart.Parameter
        End If
        task.Dispose mStart.Parameter
        Set task = Nothing
    End If
    Exit Sub
HERROR:
    mFunc.LogEx Err.Description
End Sub
