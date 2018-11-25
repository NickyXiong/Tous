VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTaskSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interface Setting"
   ClientHeight    =   6480
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9408
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTaskSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9408
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin MSComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9408
      _ExtentX        =   16595
      _ExtentY        =   508
      ButtonWidth     =   1249
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "mnuSave"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "mnuExit"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   6105
      Width           =   9405
      _ExtentX        =   16595
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4145
            MinWidth        =   4145
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgMenu 
      Left            =   2040
      Top             =   3720
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSet.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSet.frx":1C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSet.frx":202E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   1320
      Top             =   3720
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSet.frx":23C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSet.frx":2C1A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   3550
      TabIndex        =   2
      Top             =   320
      Width           =   5840
      Begin VB.Frame Frame2 
         Caption         =   "Mailbox"
         Height          =   2895
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   5415
         Begin VB.TextBox txtMail 
            Height          =   1575
            Index           =   0
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtMail 
            Height          =   1575
            Index           =   1
            Left            =   2880
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mail To:"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   16
            Top             =   480
            Width           =   555
         End
         Begin VB.Label lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CC:"
            Height          =   195
            Index           =   5
            Left            =   2880
            TabIndex        =   15
            Top             =   480
            Width           =   1500
         End
         Begin VB.Label lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter an e-mail address for each row."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   14
            Top             =   2400
            Width           =   3660
         End
      End
      Begin VB.TextBox txtTime 
         Height          =   285
         Index           =   1
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   9
         Text            =   "00:00"
         Top             =   1880
         Width           =   615
      End
      Begin VB.TextBox txtTime 
         Height          =   285
         Index           =   0
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "5"
         Top             =   800
         Width           =   615
      End
      Begin VB.OptionButton OptTime 
         Caption         =   "Daily"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton OptTime 
         Caption         =   "Real-time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(hh:mm)"
         Height          =   195
         Index           =   3
         Left            =   3000
         TabIndex        =   10
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Daily operation time"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   8
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0-120) minutes"
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   6
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operation frequency"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   840
         Width           =   1500
      End
   End
   Begin MSComctlLib.TreeView TView 
      Height          =   5685
      Left            =   0
      TabIndex        =   0
      Top             =   400
      Width           =   3500
      _ExtentX        =   6160
      _ExtentY        =   10033
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      Style           =   5
      ImageList       =   "ImgList"
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the line item set."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3600
      TabIndex        =   17
      Top             =   720
      Width           =   1680
   End
End
Attribute VB_Name = "frmTaskSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ThisIndex As Long
Private ThisTasks As TB_Context.TBTasks
Private ThisMails As TB_Context.TBMailsEx

Private Sub Form_Load()
    InitData
    InitTask
    RefreshControl
End Sub

Private Sub InitData()
    Dim o As TB_Runtime.LocalContent
    Set o = New TB_Runtime.LocalContent
    Set ThisTasks = o.GetTasks
    Set ThisMails = o.GetMails
    Set o = Nothing
End Sub

Private Sub InitTask()
    Dim Index As Long
    Dim task As TYPE_TASK
    
    With TView.Nodes
        .Add , , "TASK", "Task", 1
        For Index = 0 To ThisTasks.Size - 1
            task = ThisTasks.task(Index)
            .Add "TASK", tvwChild, task.Number, task.Description, 2
        Next Index
    End With
    With TView
        .LabelEdit = tvwManual
        .SelectedItem = TView.Nodes.Item(1)
        .Nodes.Item(1).Expanded = True
    End With
    ThisIndex = -1
End Sub

Private Sub TBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo HERROR
    Select Case Button.Key
        Case "mnuSave"
            Save
            MsgBox "Save Success!", vbInformation + vbOKOnly, mParam.CONST_RUN_TITLE
        Case "mnuExit"
            Unload Me
    End Select
    Exit Sub
HERROR:
    MsgBox Err.Description, vbExclamation + vbOKOnly, mParam.CONST_RUN_TITLE
End Sub

Private Sub Save()
    Dim o As TB_Runtime.LocalContent
    Set o = New TB_Runtime.LocalContent
    o.SetTasks ThisTasks
    o.SetMails ThisMails
    Set mRuntime.Tasks = ThisTasks
    Set o = Nothing
End Sub

Private Sub TView_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim sKey As String
    
    sKey = Node.Key
    If sKey = "TASK" Then
        ThisIndex = -1
    Else
        ThisIndex = ThisTasks.Lookup(sKey)
    End If
    
    RefreshControl
End Sub

Private Sub RefreshControl()
    Dim Index As Long, l As Long
    Dim mails As TB_Context.TBMails
    Dim task As TYPE_TASK
    Dim sMail As String
    
    If ThisIndex = -1 Then
        Frame1.Enabled = False
        Frame1.Visible = False
    Else
        Frame1.Enabled = True
        Frame1.Visible = True
        
        task = ThisTasks.task(ThisIndex)
        With task
            If .IsSys Then
                Frame2.Visible = False
            Else
                Frame2.Visible = True
            End If
            
            If .RunStyle = ENUM_RUNSTYLE.RUNSTYLE_ACTUAL Then
                OptTime(0).Value = True
            Else
                OptTime(1).Value = True
            End If
            
            txtTime(0).Text = .Interval
            txtTime(1).Text = .StartTime
            txtMail(0).Text = ""
            txtMail(1).Text = ""
            
            Index = ThisMails.Lookup(task.Number)
            If Index > -1 Then
                sMail = ""
                Set mails = ThisMails.ToMail(Index)
                If mails.Size > 0 Then
                    For l = 0 To mails.Size - 2
                        sMail = sMail & mails.Mail(l) & vbCrLf
                    Next l
                    sMail = sMail & mails.Mail(l)
                End If
                txtMail(0).Text = sMail
                
                sMail = ""
                Set mails = ThisMails.CCMail(Index)
                If mails.Size > 0 Then
                    For l = 0 To mails.Size - 2
                        sMail = sMail & mails.Mail(l) & vbCrLf
                    Next l
                    sMail = sMail & mails.Mail(l)
                End If
                txtMail(1).Text = sMail
            End If
        End With
    End If
End Sub

Private Sub OptTime_Click(Index As Integer)
    Dim tsk As TYPE_TASK
    If ThisIndex > -1 Then
        tsk = ThisTasks.task(ThisIndex)
        Select Case Index
            Case 0
                tsk.RunStyle = ENUM_RUNSTYLE.RUNSTYLE_ACTUAL
            Case 1
                tsk.RunStyle = ENUM_RUNSTYLE.RUNSTYLE_FIXED
        End Select
        ThisTasks.task(ThisIndex) = tsk
    End If
End Sub

Private Sub txtMail_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 59
            If Len(txtMail(Index).Text) > 0 Then
                If Not Asc(Right(txtMail(Index).Text, 1)) = 10 Then
                    txtMail(Index).Text = txtMail(Index).Text & vbCrLf
                End If
            End If
            txtMail(Index).SelStart = Len(txtMail(Index).Text)
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTime_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 58, 186, 8

        Case Else
            KeyAscii = 0

    End Select
End Sub

Private Sub txtTime_Validate(Index As Integer, Cancel As Boolean)
    Dim tsk As TYPE_TASK
    Dim sTime As String
On Error GoTo HERROR
    If ThisIndex > -1 Then
        sTime = txtTime(Index).Text
        tsk = ThisTasks.task(ThisIndex)
        Select Case Index
            Case 0
                If IsShortNumber(sTime) Then
                    tsk.Interval = CLng(sTime)
                Else
                    Err.Raise -1, "Validate", "Time interval is not valid data!"
                End If
            Case 1
                If IsShortTime(sTime) Then
                    tsk.StartTime = sTime
                Else
                    Err.Raise -1, "Validate", "The Daily operation time is not a valid data!"
                End If
        End Select
        txtTime(Index).Text = sTime
        ThisTasks.task(ThisIndex) = tsk
    End If
    Exit Sub
HERROR:
    If ThisIndex > -1 Then
        Select Case Index
            Case 0
                txtTime(0).Text = ThisTasks.task(ThisIndex).Interval
            Case 1
                txtTime(1).Text = ThisTasks.task(ThisIndex).StartTime
        End Select
    End If
    MsgBox Err.Description, vbExclamation + vbOKOnly, mParam.CONST_RUN_TITLE
End Sub

Private Sub txtMail_Validate(Index As Integer, Cancel As Boolean)
    Dim Index2 As Long, Index3 As Long
    Dim mails As TB_Context.TBMails
    Dim vMail() As String
    
    If ThisIndex > -1 Then
        Set mails = New TB_Context.TBMails
        vMail = Split(txtMail(Index).Text, vbCrLf)
        For Index2 = 0 To UboundEx(vMail)
            If Len(Trim(vMail(Index2))) > 0 Then
                mails.Add Trim(vMail(Index2))
            End If
        Next Index2
        Index3 = ThisMails.Lookup(ThisTasks.task(ThisIndex).Number)
        
        If Index3 > -1 Then
            Select Case Index
                Case 0
                    Set ThisMails.ToMail(Index3) = mails
                Case 1
                    Set ThisMails.CCMail(Index3) = mails
            End Select
        End If
        
        Set mails = Nothing
        Erase vMail
    End If
End Sub

Private Function IsShortNumber(sTime As String) As Boolean
    Dim l As Long
    If IsNumeric(sTime) Then
        l = CLng(sTime)
        If l >= 0 And l <= 120 Then
            sTime = l
            IsShortNumber = True
            Exit Function
        End If
    End If
    
    IsShortNumber = False
End Function

Private Function IsShortTime(sTime As String) As Boolean
    Dim v() As String
    Dim l1 As Long, l2 As Long
    
    v = Split(sTime, ":")
    If UBound(v) = 1 Then
        If IsNumeric(v(0)) And IsNumeric(v(1)) Then
            l1 = CLng(v(0))
            l2 = CLng(v(1))
            
            If l1 = 24 Then l1 = 0
            If (l1 >= 0 And l1 < 24) And (l2 >= 0 And l2 < 60) Then
                sTime = Format(l1, "00") & ":" & Format(l2, "00")
                IsShortTime = True
                Exit Function
            End If
        End If
    End If
    
    IsShortTime = False
End Function

Private Function UboundEx(v As Variant) As Long
    On Error GoTo HERROR
    UboundEx = UBound(v)
    Exit Function
HERROR:
    UboundEx = -1
End Function
