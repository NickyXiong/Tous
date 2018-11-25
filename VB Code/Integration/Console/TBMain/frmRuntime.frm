VERSION 5.00
Begin VB.Form frmRuntime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Runtime"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRuntime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Time 
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmRuntime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONST_TIME_INTERVAL As Long = 60000

Private Sub Form_Load()
    mRuntime.InitRuntime
    InitClock
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mRuntime.Dispose
End Sub

Private Sub InitClock()
    With Time
        .Interval = CONST_TIME_INTERVAL
        .Enabled = True
    End With
End Sub

Private Sub Time_Timer()
    DoAction
End Sub

Private Sub DoAction()
    Dim Index As Long
    
    For Index = 0 To Tasks.Size - 1
        If Tasks.task(Index).IsUse Then
TB_Runtime.Log "TASK", "Task " & Tasks.task(Index).Number & " Clock!"
            If IsRunTime(Tasks.task(Index).Number, 1) And Not IsRun(Tasks.task(Index).Number) Then
TB_Runtime.Log "TASK", "Task " & Tasks.task(Index).Number & " Run time."
                RunTask Tasks.task(Index).Number
            End If
        End If
    Next Index
End Sub
