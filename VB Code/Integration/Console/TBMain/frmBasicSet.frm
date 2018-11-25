VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBasicSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4644
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4812
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBasicSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   4812
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   3480
      TabIndex        =   11
      Top             =   4200
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   2040
      TabIndex        =   10
      Top             =   4200
      Width           =   1200
   End
   Begin TabDlg.SSTab Tab 
      Height          =   3972
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4572
      _ExtentX        =   8065
      _ExtentY        =   7006
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   882
      MouseIcon       =   "frmBasicSet.frx":0E42
      TabCaption(0)   =   "Database"
      TabPicture(0)   =   "frmBasicSet.frx":0E5E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "labAC(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "labAC(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "labAC(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "labAC(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAC(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmbDatabase"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAC(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAC(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdConnection"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   " FTP"
      TabPicture(1)   =   "frmBasicSet.frx":11F8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "labFTP(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "labFTP(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "labFTP(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "labFTP(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "labFTP(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtFTP(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtFTP(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtFTP(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtFTP(3)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdConnRemote"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtFTP(4)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   " SMTP"
      TabPicture(2)   =   "frmBasicSet.frx":162C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "labSMTP(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "labSMTP(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "labSMTP(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "labSMTP(3)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "labSMTP(4)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtSMTP(0)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtSMTP(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtSMTP(3)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtSMTP(4)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtSMTP(1)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "SFTP"
      TabPicture(3)   =   "frmBasicSet.frx":19C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "labFTP(5)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "labFTP(6)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "labFTP(7)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "labFTP(8)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "labFTP(9)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "labFTP(11)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtSFTP(4)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtSFTP(3)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtSFTP(0)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtSFTP(1)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "txtSFTP(2)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "txtSFTP(5)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).ControlCount=   12
      Begin VB.TextBox txtSFTP 
         Height          =   372
         Index           =   5
         Left            =   -73440
         TabIndex        =   45
         Top             =   1320
         Width           =   2532
      End
      Begin VB.TextBox txtSMTP 
         Height          =   372
         Index           =   1
         Left            =   -73440
         TabIndex        =   44
         Top             =   1320
         Width           =   2532
      End
      Begin VB.TextBox txtSFTP 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   -73440
         PasswordChar    =   "*"
         TabIndex        =   36
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtSFTP 
         Height          =   375
         Index           =   1
         Left            =   -73440
         TabIndex        =   35
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtSFTP 
         Height          =   375
         Index           =   0
         Left            =   -73440
         TabIndex        =   34
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtSFTP 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   -73440
         TabIndex        =   33
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox txtSFTP 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   -73440
         TabIndex        =   32
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox txtSMTP 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   -73440
         TabIndex        =   30
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtFTP 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   -73440
         TabIndex        =   28
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox txtSMTP 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   -73440
         TabIndex        =   26
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox txtSMTP 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   -73440
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtSMTP 
         Height          =   375
         Index           =   0
         Left            =   -73440
         TabIndex        =   21
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton cmdConnRemote 
         Caption         =   "Connect"
         Height          =   360
         Left            =   -72120
         TabIndex        =   20
         Top             =   3360
         Width           =   1248
      End
      Begin VB.TextBox txtFTP 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   -73440
         TabIndex        =   18
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtFTP 
         Height          =   375
         Index           =   0
         Left            =   -73440
         TabIndex        =   14
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtFTP 
         Height          =   375
         Index           =   1
         Left            =   -73440
         TabIndex        =   13
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtFTP 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   -73440
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton cmdConnection 
         Caption         =   "Connect"
         Height          =   360
         Left            =   2880
         TabIndex        =   5
         Top             =   2880
         Width           =   1236
      End
      Begin VB.TextBox txtAC 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtAC 
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox cmbDatabase 
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtAC 
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port："
         Height          =   204
         Index           =   11
         Left            =   -74100
         TabIndex        =   43
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password："
         Height          =   192
         Index           =   9
         Left            =   -74352
         TabIndex        =   41
         Top             =   2400
         Width           =   876
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username："
         Height          =   192
         Index           =   8
         Left            =   -74388
         TabIndex        =   40
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server："
         Height          =   192
         Index           =   7
         Left            =   -74148
         TabIndex        =   39
         Top             =   960
         Width           =   660
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Download Root："
         Height          =   192
         Index           =   6
         Left            =   -74760
         TabIndex        =   38
         Top             =   2880
         Width           =   1272
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upload Root："
         Height          =   192
         Index           =   5
         Left            =   -74556
         TabIndex        =   37
         Top             =   3360
         Width           =   1068
      End
      Begin VB.Label labSMTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port："
         Height          =   192
         Index           =   4
         Left            =   -74040
         TabIndex        =   31
         Top             =   2400
         Width           =   480
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upload Root："
         Height          =   192
         Index           =   4
         Left            =   -74628
         TabIndex        =   29
         Top             =   2880
         Width           =   1068
      End
      Begin VB.Label labSMTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sender："
         Height          =   192
         Index           =   3
         Left            =   -74244
         TabIndex        =   27
         Top             =   2880
         Width           =   696
      End
      Begin VB.Label labSMTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password："
         Height          =   192
         Index           =   2
         Left            =   -74424
         TabIndex        =   25
         Top             =   1920
         Width           =   876
      End
      Begin VB.Label labSMTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username："
         Height          =   192
         Index           =   1
         Left            =   -74460
         TabIndex        =   23
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label labSMTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server："
         Height          =   192
         Index           =   0
         Left            =   -74220
         TabIndex        =   22
         Top             =   960
         Width           =   660
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Download Root："
         Height          =   192
         Index           =   3
         Left            =   -74832
         TabIndex        =   19
         Top             =   2400
         Width           =   1272
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server："
         Height          =   192
         Index           =   0
         Left            =   -74220
         TabIndex        =   17
         Top             =   960
         Width           =   660
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username："
         Height          =   192
         Index           =   1
         Left            =   -74460
         TabIndex        =   16
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label labFTP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password："
         Height          =   192
         Index           =   2
         Left            =   -74424
         TabIndex        =   15
         Top             =   1920
         Width           =   876
      End
      Begin VB.Label labAC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database："
         Height          =   192
         Index           =   3
         Left            =   576
         TabIndex        =   9
         Top             =   2400
         Width           =   876
      End
      Begin VB.Label labAC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password："
         Height          =   192
         Index           =   2
         Left            =   576
         TabIndex        =   8
         Top             =   1920
         Width           =   876
      End
      Begin VB.Label labAC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username："
         Height          =   192
         Index           =   1
         Left            =   540
         TabIndex        =   7
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label labAC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server："
         Height          =   192
         Index           =   0
         Left            =   780
         TabIndex        =   6
         Top             =   960
         Width           =   660
      End
   End
   Begin VB.Label labFTP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username："
      Height          =   192
      Index           =   10
      Left            =   720
      TabIndex        =   42
      Top             =   1560
      Width           =   900
   End
End
Attribute VB_Name = "frmBasicSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    LoadSetting
End Sub

Private Sub cmdSave_Click()
On Error GoTo HERROR

    CheckSetting
    SaveSetting
    Unload Me
    Exit Sub
HERROR:
    MsgBox Err.Description, vbExclamation + vbOKOnly, mParam.CONST_RUN_TITLE
End Sub

Private Sub CheckSetting()
    If Len(Trim(txtAC(0).Text)) = 0 Then
        Err.Raise -1, "CheckSetting", "must input A/C server address！"
    End If
    If Len(Trim(txtAC(0).Text)) = 0 Then
        Err.Raise -1, "CheckSetting", "must input A/C username！"
    End If
    If Len(Trim(cmbDatabase.Text)) = 0 Then
        Err.Raise -1, "CheckSetting", "connection and select database！"
    End If
    
    If Len(Trim(txtFTP(0).Text)) = 0 Then
        Err.Raise -1, "CheckSetting", "must input FTP server address！"
    End If
    If Len(Trim(txtFTP(1).Text)) = 0 Then
        Err.Raise -1, "CheckSetting", "must input FTP username！"
    End If
    
    If Len(Trim(txtSMTP(0).Text)) = 0 Then
        Err.Raise -1, "CheckSetting", "must input SMTP server address！"
    End If
'    If Len(Trim(txtSMTP(1).Text)) = 0 Then
'        Err.Raise -1, "CheckSetting", "must input SMTP username!"
'    End If
    If Len(Trim(txtSMTP(3).Text)) = 0 Then
        Err.Raise -1, "CheckSetting", "must input sender！"
    End If
End Sub

Private Sub LoadSetting()
    Dim Index As Long
    Dim o As TB_Runtime.LocalContent
    Set o = New TB_Runtime.LocalContent
    
    Dim strReturn As String
    strReturn = String(100, 0)
    
    Dim k3svr As TYPE_K3SERVER
    k3svr = o.GetK3Server
    
    With k3svr
        txtAC(0).Text = .DBServer
        cmbDatabase.Text = .DBName
        txtAC(1).Text = .DBUsername
        txtAC(2).Text = .DBPassword
    End With
    
    Dim rmts As TB_Context.TBRemotes
    Dim rmt As TYPE_REMOTE
    Set rmts = o.GetRemotes
    Index = rmts.Lookup(ENUM_FTPROOT.ROOT_TBHQ Or ENUM_FTPROOT.ROOT_3PL Or ENUM_FTPROOT.ROOT_POS)
    If Index > -1 Then
        rmt = rmts.Remote(Index)
        
        With rmt
            txtFTP(0).Text = .Server
            txtFTP(1).Text = .Username
            txtFTP(2).Text = .Password
            txtFTP(3).Text = .DownRoot
            txtFTP(4).Text = .UpRoot
        End With
    End If
    Set rmts = Nothing
    
    Dim smp As TYPE_EMAILSMTP
    smp = o.GetSmtp
    With smp
        txtSMTP(0).Text = .Smtp
        txtSMTP(1).Text = .Username
        txtSMTP(2).Text = .Password
        txtSMTP(4).Text = .Port
        txtSMTP(3).Text = .Sender
    End With
    
    Set o = Nothing
    
    '读取SFTP配置
    GetPrivateProfileString "SFTP", "Server", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    txtSFTP(0).Text = Replace(strReturn, Chr(0), "")
    
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "Username", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    txtSFTP(1).Text = Replace(strReturn, Chr(0), "")
    
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "Password", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    txtSFTP(2).Text = Replace(strReturn, Chr(0), "")
    
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "DownloadRoot", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    txtSFTP(3).Text = Replace(strReturn, Chr(0), "")
    
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "UploadRoot", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    txtSFTP(4).Text = Replace(strReturn, Chr(0), "")
    
    strReturn = String(100, 0)
    GetPrivateProfileString "SFTP", "Port", "", strReturn, 100, App.path & "\Setting\SFTP.ini"
    txtSFTP(5).Text = Replace(strReturn, Chr(0), "")
    
End Sub

Private Sub SaveSetting()
    Dim Index As Long
    
    Dim o As TB_Runtime.LocalContent
    Set o = New TB_Runtime.LocalContent
    
    Dim k3svr As TYPE_K3SERVER
    With k3svr
        .DBServer = txtAC(0).Text
        .DBName = cmbDatabase.Text
        .DBUsername = txtAC(1).Text
        .DBPassword = txtAC(2).Text
    End With
    o.SetK3Server k3svr
    
    Dim rmts As TB_Context.TBRemotes
    Dim rmt As TYPE_REMOTE
    Set rmts = o.GetRemotes
    Index = rmts.Lookup(ENUM_FTPROOT.ROOT_TBHQ Or ENUM_FTPROOT.ROOT_3PL Or ENUM_FTPROOT.ROOT_POS)
    If Index > -1 Then
        rmt = rmts.Remote(Index)
        
        With rmt
            .Server = txtFTP(0).Text
            .Username = txtFTP(1).Text
            .Password = txtFTP(2).Text
            .DownRoot = txtFTP(3).Text
            .UpRoot = txtFTP(4).Text
        End With
    Else
        With rmt
            .Server = txtFTP(0).Text
            .Username = txtFTP(1).Text
            .Password = txtFTP(2).Text
            .DownRoot = txtFTP(3).Text
            .UpRoot = txtFTP(4).Text
            .BackupRoot = "Backup\"
            .CacheRoot = "Cache\"
            .Name = "TBHQ|3PL|POS"
            .RangeID = ENUM_FTPRANGE.RANGE_UNION
            .RootID = ENUM_FTPROOT.ROOT_TBHQ Or ENUM_FTPROOT.ROOT_3PL Or ENUM_FTPROOT.ROOT_POS
        End With
    End If
    rmts.Add rmt
    o.SetRemotes rmts
    Set rmts = Nothing
    
    Dim smp As TYPE_EMAILSMTP
    With smp
        .Smtp = txtSMTP(0).Text
        .Username = txtSMTP(1).Text
        .Password = txtSMTP(2).Text
        .Port = CInt(Val(txtSMTP(4).Text))
        .Sender = txtSMTP(3).Text
    End With
    o.SetSmtp smp
    
    Set o = Nothing
    
    '写入SFTP配置
    WritePrivateProfileString "SFTP", "Server", txtSFTP(0).Text, App.path & "\Setting\SFTP.ini"
    WritePrivateProfileString "SFTP", "Username", txtSFTP(1).Text, App.path & "\Setting\SFTP.ini"
    WritePrivateProfileString "SFTP", "Password", txtSFTP(2).Text, App.path & "\Setting\SFTP.ini"
    WritePrivateProfileString "SFTP", "DownloadRoot", txtSFTP(3).Text, App.path & "\Setting\SFTP.ini"
    WritePrivateProfileString "SFTP", "UploadRoot", txtSFTP(4).Text, App.path & "\Setting\SFTP.ini"
    WritePrivateProfileString "SFTP", "Port", txtSFTP(5).Text, App.path & "\Setting\SFTP.ini"
    
End Sub

Private Sub cmdConnection_Click()
    Dim lIndex As Long
    Dim bRet As Boolean
    Dim sStr As String
    Dim arr() As String

    cmbDatabase.Clear
    
    bRet = TryConnection(txtAC(0), txtAC(1), txtAC(2), cmbDatabase.Text, sStr)
    If Not bRet Then
        MsgBox "connection fail：" & sStr, vbExclamation + vbOKOnly, mParam.CONST_RUN_TITLE
    Else
        arr = GetConnectionDatabase(txtAC(0), txtAC(1), txtAC(2))
        For lIndex = 1 To UBound(arr)
            cmbDatabase.AddItem arr(lIndex)
        Next lIndex
        MsgBox "connection success！", vbInformation + vbOKOnly, mParam.CONST_RUN_TITLE
    End If
    Erase arr
End Sub

Private Sub cmdConnRemote_Click()
    Dim oInternet As XZInternet
    Set oInternet = New XZInternet
    
    If oInternet.Connection(txtFTP(0).Text, txtFTP(1).Text, txtFTP(2).Text) Then
        MsgBox "connection success！", vbInformation + vbOKOnly, mParam.CONST_RUN_TITLE
    Else
        MsgBox "connection failed!", vbExclamation + vbOKOnly, mParam.CONST_RUN_TITLE
    End If
    
    oInternet.Dispose
    Set oInternet = Nothing
End Sub

Private Sub cmdConnSMTP_Click()
    Dim o As Object
    Set o = CreateObject("jmail.Message")
    Set o = Nothing
End Sub

Public Function TryConnection(ByVal sServer As String, ByVal sUsername As String, ByVal sPasswrod As String, ByVal sDatabase As String, _
                            Optional ByRef sResult As String, Optional ByRef bConnectionDatabse As Boolean = False) As Boolean
    Dim cnString As String
    Dim cn As ADODB.Connection
    
On Error GoTo HERROR

    If bConnectionDatabse Then
        cnString = "Provider=SQLOLEDB.1;User ID=" & sUsername & ";Password=" & sPasswrod & ";Data Source=" & sServer
    Else
        cnString = "Provider=SQLOLEDB.1;User ID=" & sUsername & ";Password=" & sPasswrod & ";Data Source=" & sServer & ";Initial Catalog=" & sDatabase
    End If
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = cnString
    cn.ConnectionTimeout = 5
    cn.CursorLocation = adUseClient
    cn.Open
    
    cn.Close
    Set cn = Nothing
    TryConnection = True
    Exit Function
HERROR:
    sResult = Err.Description
    TryConnection = False
    Set cn = Nothing
End Function

Public Function GetConnectionDatabase(ByVal sServer As String, ByVal sUsername As String, ByVal sPasswrod As String) As String()
    Dim cnString As String
    Dim cn As ADODB.Connection
    Dim lIndex As Long
    Dim arr() As String
    Dim rs As ADODB.Recordset
On Error GoTo HERROR

    cnString = "Provider=SQLOLEDB.1;User ID=" & sUsername & ";Password=" & sPasswrod & ";Data Source=" & sServer
    Set cn = New ADODB.Connection
    cn.ConnectionString = cnString
    cn.ConnectionTimeout = 5
    cn.CursorLocation = adUseClient
    cn.Open
    
    Set rs = cn.Execute("SELECT NAME FROM sys.databases")
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            ReDim arr(rs.RecordCount)
            For lIndex = 1 To rs.RecordCount
                arr(lIndex) = rs("NAME")
                rs.MoveNext
            Next lIndex
        End If
    End If
    GetConnectionDatabase = arr
    Erase arr
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    Exit Function
HERROR:
    Set rs = Nothing
    Set cn = Nothing
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

