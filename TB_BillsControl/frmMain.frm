VERSION 5.00
Object = "{CEA2DF91-E53D-11D5-9FA9-00E04C54B3B6}#1.25#0"; "K3List.ocx"
Begin VB.Form frmMain 
   Caption         =   "主界面"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   10020
   WindowState     =   2  'Maximized
   Begin K3List.ICList ICList 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11033
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ok As Boolean

Public Property Get Selected() As KFO.Vector
    Set Selected = ICList.GetSelected()
End Property

Public Property Get OK() As Boolean
    OK = m_ok
End Property

Private Sub Form_Load()
    With ICList
        .Left = 0
        .Top = 0
    End With
    m_ok = False
End Sub

Private Sub Form_Resize()
    If Me.ScaleHeight > 0 Then
        ICList.Height = Me.ScaleHeight
    End If
    If Me.ScaleWidth > 0 Then
        ICList.Width = Me.ScaleWidth
    End If
End Sub

Private Sub ICList_LogicEvents(ActionName As String, ExtInfo As String)
    If ActionName = "SelectOK" Then
        If ICList.OK Then
            m_ok = True
            Me.Visible = False
        End If
    ElseIf ActionName = "Exit" Then
        Me.Visible = False
        m_ok = False
    End If
End Sub

