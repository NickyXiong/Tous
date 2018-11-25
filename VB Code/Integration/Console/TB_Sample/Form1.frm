VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Frist"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Last"
      Height          =   360
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Perv"
      Height          =   360
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   990
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Frist"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   720
      Width           =   10695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   240
      Width           =   7125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ThisNoteset As Noteset



Private Sub Command4_Click()
    Dim v(2) As Long
    v(0) = 3
    v(1) = 7
    v(2) = 10
    Set ThisNoteset = New Noteset
    ThisNoteset.Loaddata "C:\RM_PSSALE_42.txt", 7
    'ThisNoteset.Loaddata "C:\RM_PSSALE_42.txt", v
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            ThisNoteset.IndexFirst
        Case 1
            If Not ThisNoteset.IndexNext Then
                ThisNoteset.IndexLast
            End If
        Case 2
            If Not ThisNoteset.IndexPrevious Then
                ThisNoteset.IndexFirst
            End If
        Case 3
            ThisNoteset.IndexLast
        Case Else
            Exit Sub
    End Select
    
    InitData
End Sub

Private Sub InitData()
    Dim v1 As Variant
    Dim v2 As Variant
    
    Dim l1 As Long, l2 As Long
    Dim sText As String
    
    'Label1.Caption = ThisNoteset.GetKeyByIndex
    v1 = ThisNoteset.GetDataByIndex
    
    For l1 = 0 To UBound(v1)
        v2 = v1(l1)
        For l2 = 0 To UBound(v2)
            sText = sText & v2(l2) & vbTab
        Next l2
        
        sText = Left(sText, Len(sText) - 1)
        sText = sText & vbCrLf
    Next l1
    
    Text1.Text = sText
    DoEvents
End Sub

