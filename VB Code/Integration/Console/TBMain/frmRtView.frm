VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRtView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Runtime View"
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
   Icon            =   "frmRtView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9408
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Time 
      Left            =   600
      Top             =   3840
   End
   Begin MSComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   0
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
            Caption         =   "Run"
            Key             =   "mnuRun"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "mnuStop"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6105
      Width           =   9405
      _ExtentX        =   16595
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
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
   Begin MSComctlLib.ListView lsView 
      Height          =   5685
      Left            =   3520
      TabIndex        =   2
      Top             =   400
      Width           =   5890
      _ExtentX        =   10393
      _ExtentY        =   10033
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ImgMenu"
      SmallIcons      =   "ImgMenu"
      ColHdrIcons     =   "ImgMenu"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtView.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtView.frx":11DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtView.frx":1576
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtView.frx":1910
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtView.frx":1CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtView.frx":24FC
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtView.frx":2D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtView.frx":35A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtView.frx":3DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtView.frx":4644
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TView 
      Height          =   5685
      Left            =   0
      TabIndex        =   1
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmRtView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ThisIndex As Long

Private Sub Form_Load()
    InitView
    InitTask
    InitClock
End Sub

Private Sub Form_Activate()
    RefreshRuntime
End Sub

Private Sub InitTask()
    Dim Index As Long
    Dim task As TYPE_TASK
    
    With TView.Nodes
        .Add , , "TASK", "Task", 1
        For Index = 0 To mRuntime.Tasks.Size - 1
            task = mRuntime.Tasks.task(Index)
            .Add "TASK", tvwChild, task.Number, task.Description, 2
        Next Index
    End With
    With TView
        .LabelEdit = tvwManual
        .SelectedItem = TView.Nodes.Item(1)
        .Nodes.Item(1).Expanded = True
    End With
    ThisIndex = -1
    RefreshControl
End Sub

Private Sub InitClock()
    With Time
        .Interval = 5000
        .Enabled = True
    End With
End Sub

Private Sub InitView()
    Dim Index As Long
    
    With lsView
        .View = lvwReport
        .MultiSelect = True
        .FullRowSelect = False
        .LabelEdit = lvwManual
        .ListItems.Clear
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, "Filename", "Name", 2500
        .ColumnHeaders.Add 2, "Filesize", "Size", 1500
    End With
End Sub

Private Sub lsView_DblClick()
    If Not lsView.SelectedItem Is Nothing Then
        ShellOpen lsView.SelectedItem.Key
    End If
End Sub

Private Sub lsView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lsView.HitTest(X, Y) Is lsView.SelectedItem Then
        Set lsView.SelectedItem = Nothing
    End If
End Sub

Private Sub TBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "mnuRun"
            DoRun
        Case "mnuStop"
            DoStop
        Case "mnuExit"
            Unload Me
    End Select
End Sub

Private Sub Time_Timer()
    RefreshRuntime
End Sub

Private Sub TView_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim sKey As String
    
    sKey = Node.Key
    If sKey = "TASK" Then
        ThisIndex = -1
    Else
        ThisIndex = mRuntime.Tasks.Lookup(sKey)
    End If
    
    RefreshControl
End Sub

Private Sub RefreshControl()
    Dim Index As Long, imgIndex As Long
    Dim TaskNumber As String
    Dim locDir As String, posix As String
    
    Dim oSys As FileSystemObject
    Dim oFile As File, oFolder As Folder
    
    If ThisIndex = -1 Then
        TaskNumber = "TASK"
    Else
        TaskNumber = mRuntime.Tasks.task(ThisIndex).Number
    End If
    
    locDir = TB_Runtime.DirectoryAsLog(TaskNumber)
    TB_Runtime.MakeDirByLocal locDir
    
    Set oSys = New FileSystemObject
    Set oFolder = oSys.GetFolder(locDir)
    
    With lsView
        .ListItems.Clear
        
        If oFolder.Files.Count > 0 Then
            For Each oFile In oFolder.Files
                
                Index = Index + 1
                posix = RetPosix(oFile.Name)
                
                Select Case posix
                    Case "HTM", "HTML"
                        imgIndex = 5
                    Case "LOG"
                        imgIndex = 6
                    Case Else
                        imgIndex = 1
                End Select
                
                .ListItems.Add , "Key" & Index, oFile.Name, imgIndex, imgIndex
                .ListItems("Key" & Index).ListSubItems.Add 1, , IIf((oFile.Size Mod 1024) / 1024 > 0, oFile.Size \ 1024 + 1, oFile.Size / 1024) & " KB"
                Set oFile = Nothing
            Next
        End If
    End With
    
    Set oFolder = Nothing
    Set oSys = Nothing
End Sub

Private Sub ShellOpen(ByVal itmKey As String)
    Dim TaskNumber As String
    Dim locDir As String
    Dim posix As String
    
    If ThisIndex = -1 Then
        TaskNumber = "TASK"
    Else
        TaskNumber = mRuntime.Tasks.task(ThisIndex).Number
    End If
    locDir = TB_Runtime.DirectoryAsLog(TaskNumber) & lsView.ListItems(itmKey).Text
    
    If Len(Dir(locDir, vbArchive)) > 0 Then
        posix = RetPosix(locDir)
        Select Case posix
            Case "LOG"
                Shell "notepad.exe " & locDir, vbNormalFocus
            Case Else
                mParam.ShellOpen Me.hWnd, locDir
        End Select
    End If
End Sub

Private Sub RefreshRuntime()
    Dim Index As Long
    
    With TView.Nodes
        For Index = 2 To .Count
            If mRuntime.IsRun(.Item(Index).Key) Then
                TView.Nodes.Item(Index).Image = 3
            Else
                TView.Nodes.Item(Index).Image = 4
            End If
        Next Index
    End With
End Sub

Private Sub DoRun()
    DoRunning True
End Sub

Private Sub DoStop()
    DoRunning False
End Sub

Private Sub DoRunning(Optional ByVal bRun As Boolean = True)
    Dim Index As Long
    
    If TView.SelectedItem.Key = "TASK" Then
        For Index = 0 To mRuntime.Tasks.Size - 1
            If mRuntime.IsRun(mRuntime.Tasks.task(Index).Number) <> bRun Then
                If bRun Then
                    mRuntime.RunTask mRuntime.Tasks.task(Index).Number
                Else
                    mRuntime.StopTask mRuntime.Tasks.task(Index).Number
                End If
            End If
        Next Index
    Else
        If mRuntime.IsRun(TView.SelectedItem.Key) <> bRun Then
            If bRun Then
                mRuntime.RunTask TView.SelectedItem.Key
            Else
                mRuntime.StopTask TView.SelectedItem.Key
            End If
        End If
    End If
End Sub

Private Function RetPosix(ByVal locDir As String) As String
    RetPosix = UCase(Right(locDir, Len(locDir) - InStr(1, locDir, ".")))
End Function
