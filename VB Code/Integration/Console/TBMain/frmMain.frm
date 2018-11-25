VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2544
   ClientLeft      =   120
   ClientTop       =   708
   ClientWidth     =   3672
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2544
   ScaleWidth      =   3672
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   0
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":07DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMainMenu 
      Caption         =   "MainMenu"
      Begin VB.Menu mnuBasicSet 
         Caption         =   "Setting"
      End
      Begin VB.Menu mnuTaskSet 
         Caption         =   "Interface Config"
      End
      Begin VB.Menu mnuConfigLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRtView 
         Caption         =   "Runtime View"
      End
      Begin VB.Menu mnuConfieLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    InitControl
End Sub

Private Sub InitControl()
    mTray.LoadIcon True
    Me.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sMessage As Long
    
    sMessage = X / Screen.TwipsPerPixelX

    Select Case sMessage
        
        Case WM_RBUTTONUP
        
            SetForegroundWindow Me.hWnd
            Me.PopupMenu mnuMainMenu
        
        Case WM_LBUTTONDBLCLK
            
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mTray.UnLoadIcon
End Sub

Private Sub mnuBasicSet_Click()
    frmBasicSet.Show
End Sub

Private Sub mnuTaskSet_Click()
    frmTaskSet.Show
End Sub

Private Sub mnuRtView_Click()
    frmRtView.Show
End Sub

Private Sub mnuExit_Click()
    mStart.Dispose
End Sub
