VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgBar 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleMode       =   0  'User
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblMsg 
      Height          =   480
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents m_clsPass As clsPass
Attribute m_clsPass.VB_VarHelpID = -1
Public Event Active()

Private Sub Form_Activate()
    m_clsPass.SetActive
End Sub


Private Sub Form_Load()
    On Error Resume Next
   ' SetFormFont Me
    'Animation1.Open App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & "KDBook.ANI"
    Err.Clear
End Sub

Private Sub m_clsPass_HideProgBar()
    ProgressBar1.Visible = False
    DoEvents
End Sub

Private Sub m_clsPass_SetBarMaxValue(ByVal Value As Long)
    ProgressBar1.Max = Value + 1
    DoEvents
End Sub

Private Sub m_clsPass_SetBarMinValue(ByVal Value As Long)
    ProgressBar1.Min = Value
    DoEvents
End Sub

Private Sub m_clsPass_SetBarValue(ByVal Value As Long)
    ProgressBar1.Value = Value
    DoEvents
End Sub

Private Sub m_clsPass_SetBarValueWithMax()
    ProgressBar1.Value = ProgressBar1.Max
    DoEvents
End Sub

Private Sub m_clsPass_SetMsg(ByVal Msg As String)
    lblMsg.Caption = Msg
    DoEvents
End Sub

Private Sub m_clsPass_ShowProgBar()
    ProgressBar1.Visible = True
    DoEvents
End Sub

Private Sub m_clsPass_Unload()
    ProgressBar1.Value = ProgressBar1.Max
    DoEvents
    Unload Me
End Sub

