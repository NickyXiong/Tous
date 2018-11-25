VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PF - AR Invoice & Voucher Importation"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "FTP Setting"
      TabPicture(0)   =   "frmMain.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Inet1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btnSave"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnDownloadFile"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Auto Importation"
      TabPicture(1)   =   "frmMain.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblStatus"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblAccountName"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "CmdCancelImport"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "CmdOK"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Manual Importation"
      TabPicture(2)   =   "frmMain.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmMain.frx":0E96
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CommonDialog1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmMain.frx":0EB2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "CommonDialog2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2220
         Left            =   -74880
         TabIndex        =   25
         Top             =   1320
         Width           =   6735
         Begin VB.TextBox txtFailure 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            TabIndex        =   28
            Top             =   1200
            Width           =   4800
         End
         Begin VB.TextBox txtNotImported 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            TabIndex        =   27
            Top             =   240
            Width           =   4800
         End
         Begin VB.Timer tmrUpdate 
            Left            =   3480
            Top             =   1680
         End
         Begin VB.TextBox txtImported 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            TabIndex        =   26
            Top             =   720
            Width           =   4800
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   1800
            TabIndex        =   29
            Top             =   1800
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   582
            _Version        =   393216
            Format          =   60751874
            CurrentDate     =   40238.3333333333
         End
         Begin VB.Label Label5 
            Caption         =   "Failure Folder:"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Not Imported Folder:"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Daily Importation Time"
            Height          =   225
            Left            =   120
            TabIndex        =   31
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Imported Folder:"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "Activate automatic import"
         Height          =   375
         Left            =   -72720
         TabIndex        =   24
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CommandButton CmdCancelImport 
         Caption         =   "Disable automatic import"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70320
         TabIndex        =   23
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Height          =   2580
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   6735
         Begin VB.TextBox txtUserName 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1080
            TabIndex        =   18
            Top             =   960
            Width           =   5520
         End
         Begin VB.TextBox txtFtpServerUrl 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1080
            TabIndex        =   17
            Top             =   240
            Width           =   5520
         End
         Begin VB.TextBox txtPwd 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1080
            PasswordChar    =   "*"
            TabIndex        =   16
            Top             =   1320
            Width           =   5520
         End
         Begin VB.CommandButton btnTestConn 
            Caption         =   "Test Connection"
            Height          =   375
            Left            =   4440
            TabIndex        =   15
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox txtFtpFolder 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            Top             =   600
            Width           =   5535
         End
         Begin VB.Label Label6 
            Caption         =   "User Name"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "FTP Folder"
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Password"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Ftp folder"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.CommandButton btnDownloadFile 
         Caption         =   "Download file manually"
         Height          =   375
         Left            =   4680
         TabIndex        =   12
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save Connection"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Caption         =   "AR Invoice importation"
         Height          =   1620
         Left            =   -74880
         TabIndex        =   6
         Top             =   1080
         Width           =   6735
         Begin VB.CommandButton cmdImportInvoice 
            Caption         =   "Import"
            Height          =   375
            Left            =   5400
            TabIndex        =   9
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtInvoiceFilePath 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            TabIndex        =   8
            Top             =   360
            Width           =   5280
         End
         Begin VB.CommandButton CmdSelectInvoice 
            Height          =   300
            Left            =   6240
            Picture         =   "frmMain.frx":0ECE
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   360
            Width           =   350
         End
         Begin VB.Label Label11 
            Caption         =   "File Path"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Voucher importation"
         Height          =   1620
         Left            =   -74880
         TabIndex        =   1
         Top             =   2760
         Width           =   6735
         Begin VB.TextBox txtVoucherFilePath 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            TabIndex        =   4
            Top             =   360
            Width           =   5280
         End
         Begin VB.CommandButton cmdImportVoucher 
            Caption         =   "Import"
            Height          =   375
            Left            =   5400
            TabIndex        =   3
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdSelectVoucher 
            Height          =   300
            Left            =   6240
            Picture         =   "frmMain.frx":11A4
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   360
            Width           =   350
         End
         Begin MSComDlg.CommonDialog cmdlg 
            Left            =   1680
            Top             =   960
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label7 
            Caption         =   "File Path"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   -74640
         Top             =   5400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -65520
         Top             =   7740
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         InitDir         =   "D:\\"
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   840
         Top             =   3840
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         Protocol        =   2
         RemotePort      =   21
         URL             =   "ftp://"
      End
      Begin VB.Label Label3 
         Caption         =   " Account Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblAccountName 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -73440
         TabIndex        =   35
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label lblStatus 
         Caption         =   " Auto Importation has been cancelled"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   3600
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NICONDATA As NOTIFYICONDATA    ' 系统托盘图标结构
Dim BMenuMainW As Boolean          ' 主菜单是否显示
Dim strFolder As String
Dim iExit As Long

'***************   Added by Nicky   ****************
Dim dtImportedDate As Date          ' 执行过的日期，若当天已经执行过导入，则不再执行
Dim bIsImported As Boolean          ' 当天是否执行过Import
Dim bIsImporting As Boolean         ' 是否正在执行Import
Dim bIsExecuteOnce As Boolean       ' 是否仅执行一次
Dim bIsFirstStepOK As Boolean       ' 第一步是否成功

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private WithEvents m_ProgBar As Bar
Attribute m_ProgBar.VB_VarHelpID = -1
Private m_WaitType As Long '1---import to excel
Private m_strFileName As String '当前正在导入的文件的文件名,仅在导入时有效
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private m_lngPOSTranTypeID As Long
Private m_lngPOSSubTypeID As Long
Private lUserID As Long
Private strDocument As String
Private OrigAmount As Double
Private strUUID As String


Private shlShell As Shell32.Shell
Private shlFolder As Shell32.Folder
Private Const BIF_RETURNONLYFSDIRS = &H1


Private logWriter As New clsTextLog

Private Sub btnDownloadFile_Click()
    On Error GoTo EHandler
    DownloadFilesFromFtp txtNotImported.Text
    MsgBox "Data files have been downloaded!", vbInformation, "Kingdee Prompt"
    Exit Sub
EHandler:
    MsgBox Err.Description, vbCritical, "Kingdee Prompt"
End Sub


Private Sub btnSave_Click()
    On Error GoTo EHandler
    Call SaveFtpSettings
    MsgBox "FTP information has been saved!", vbInformation, "Kingdee Prompt"
    Exit Sub
EHandler:
    MsgBox Err.Description, vbCritical, "Kingdee Prompt"
End Sub


Private Sub ReadFtpSettings()
On Error GoTo EHandler
    Dim rs As ADODB.Recordset
    Set rs = ExecSql("Select * From PF_t_FtpSetting Where FKey='FTP'")
    If Not rs Is Nothing Then
        If rs.RecordCount = 1 Then
            txtFtpServerUrl.Text = rs.Fields("FFtpUrl").Value
            txtFtpFolder.Text = rs.Fields("FDir").Value
            txtUserName.Text = rs.Fields("FUserName").Value
            txtPwd.Text = rs.Fields("FP").Value
        End If
    Else
        MsgBox "Please apply the db script of k3 to the middle server!", vbCritical, "Kingdee Prompt..."
    End If
    Exit Sub
EHandler:
    MsgBox "Read ftp settings error:" & Err.Description, vbCritical, "Kingdee Prompt..."
End Sub

Private Sub SaveFtpSettings()
    Dim sql As New StringBuilder
    sql.Append "Delete From PF_t_FtpSetting Where FKey='FTP'" & vbCrLf
    sql.Append "Insert into PF_t_FtpSetting(FKey,FFtpUrl,FDir ,FUserName,FP) "
    sql.Append "values('FTP','"
    sql.Append SafetyStr(txtFtpServerUrl.Text) & "','"
    sql.Append SafetyStr(txtFtpFolder.Text) & "','"
    sql.Append SafetyStr(txtUserName.Text) & "','"
    sql.Append SafetyStr(txtPwd.Text) & "')"
    
    ExecSql sql.StringValue
End Sub

Private Sub btnTestConn_Click()
On Error GoTo EHandler
    Inet1.URL = txtFtpServerUrl.Text
    Inet1.UserName = txtUserName.Text
    Inet1.Password = txtPwd.Text
    SafetyExecFtpCmd "Dir"
    MsgBox "Connection is set up successfully!"
    Exit Sub
EHandler:
    MsgBox Err.Description, vbCritical, "Kingdee Prompt"
End Sub

'从Ftp站点下载文件
Private Sub DownloadFilesFromFtp(path As String)
On Error GoTo EHandler
    Inet1.URL = txtFtpServerUrl.Text
    Inet1.UserName = txtUserName.Text
    Inet1.Password = txtPwd.Text
    
    SafetyExecFtpCmd "CD " & txtFtpFolder.Text
    SafetyExecFtpCmd "Dir"
    
    While Inet1.StillExecuting
        DoEvents
    Wend
    Dim ftpfiles As String
    ftpfiles = Trim(Inet1.GetChunk(1024, icString))
    
    Dim files() As String
    files = Split(ftpfiles, vbCrLf)
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim i As Long
    For i = LBound(files) To UBound(files)
        If files(i) <> "" Then
            'If the file has been existed in the local folder,then remove it
            If fso.FileExists(path & "\" & files(i)) = True Then
                fso.DeleteFile path & "\" & files(i)
            End If
            SafetyExecFtpCmd "Get " & files(i) & " " & path & "\" & files(i)
            SafetyExecFtpCmd "Delete " & files(i)
        End If
    Next
    While Inet1.StillExecuting
        DoEvents
    Wend
EHandler:
    
End Sub

Private Sub SafetyExecFtpCmd(cmd As String)
    While Inet1.StillExecuting
        DoEvents
    Wend
    Inet1.Execute , cmd
End Sub


Private Sub cmdImportInvoice_Click()
On Error GoTo EHandler
    If txtInvoiceFilePath.Text <> "" Then
        '检查文件是否存在
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FileExists(txtInvoiceFilePath.Text) Then
            Set fso = Nothing
            Exit Sub
        End If
        Set fso = Nothing
        ImportExcelWithProgress txtInvoiceFilePath.Text, True
        
        If bIsFirstStepOK = True Then
            MoveFile txtInvoiceFilePath.Text, txtImported.Text & "\" & GetFileNameWithoutPath(txtInvoiceFilePath.Text)
            logWriter.DeleteFile
            MsgBox "The file has been imported!", vbInformation, "Kingdee Prompt"
        Else
            MoveFile txtInvoiceFilePath.Text, txtFailure.Text & "\" & GetFileNameWithoutPath(txtInvoiceFilePath.Text)
            MsgBox "The file is failed to import!", vbInformation, "Kingdee Prompt"
        End If
        
    Else
        MsgBox "Please select a file and try again!", vbOKOnly, "Kingdee Prompt..."
    End If
    
    Exit Sub
EHandler:
    MsgBox Err.Description, vbCritical, "Kingdee Prompt..."
End Sub

Private Sub cmdImportVoucher_Click()
    On Error GoTo EHandler
    
    If txtVoucherFilePath.Text <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FileExists(txtVoucherFilePath.Text) Then
            Set fso = Nothing
            Exit Sub
        End If
        Set fso = Nothing
        ImportExcelWithProgress txtVoucherFilePath.Text, True
        
        If bIsFirstStepOK = True Then
            MoveFile txtVoucherFilePath.Text, txtImported.Text & "\" & GetFileNameWithoutPath(txtVoucherFilePath.Text)
            logWriter.DeleteFile
            MsgBox "The file has been imported!", vbInformation, "Kingdee Prompt"
        Else
            MoveFile txtVoucherFilePath.Text, txtFailure.Text & "\" & GetFileNameWithoutPath(txtVoucherFilePath.Text)
            MsgBox "The file imported failed!", vbInformation, "Kingdee Prompt"
        End If
    Else
        MsgBox "Please select a file and try again!", vbOKOnly, "Kingdee Prompt..."
    End If
    Exit Sub
EHandler:
    MsgBox Err.Description, vbCritical, "Kingdee Prompt..."
End Sub

Private Sub CmdOK_Click()
    Dim response
    'response = MsgBox("Whether you want to execute the automation once immediately?", vbYesNoCancel, "Kingdee Prompt")
    response = vbNo
    If response = vbYes Then

        lblStatus.Caption = "Auto Importation has already been started"
        CmdOK.Enabled = False
        CmdCancelImport.Enabled = True
        frmMain.Refresh
        bIsExecuteOnce = True
        tmrUpdate.Interval = 500
        tmrUpdate.Enabled = True
    ElseIf response = vbNo Then

        lblStatus.Caption = "Auto Importation has already been started"
        CmdOK.Enabled = False
        CmdCancelImport.Enabled = True
        frmMain.Refresh
        bIsExecuteOnce = False
        tmrUpdate.Interval = 20000
        tmrUpdate.Enabled = True
    Else
        Exit Sub
    End If
End Sub

Private Sub CmdCancelImport_Click()
    tmrUpdate.Enabled = False
    lblStatus.Caption = "Auto Importation has been cancelled"
    CmdOK.Enabled = True
    CmdCancelImport.Enabled = False
    frmMain.Refresh
End Sub

'Edit by Nicky 2010-3-3
'导入指定源路径下的csv文件
Private Sub ImportFiles(ByVal strSourcePath As String, ByVal strSucPath As String, ByVal strFailPath As String)
    Dim strfilename As String
    Dim arrFileList() As String '用于存放需要导入的文件名称
    Dim i As Integer
    If Len(strSourcePath) <> 0 Then
        '读取导入目录下的所有csv文件列表
        i = 0
        ReDim Preserve arrFileList(i) As String
       
        strfilename = Dir(strSourcePath & "\*.csv")
        Do While strfilename <> ""
            If UCase(Right(strfilename, 3)) = "CSV" Then
                i = i + 1
                ReDim Preserve arrFileList(i) As String
                arrFileList(i) = strfilename
            End If
            strfilename = Dir() '读取下一个文件
        Loop
            
        If UBound(arrFileList) = 0 Then
            Exit Sub
        End If
        
        '开始逐文件导入
        For i = 1 To UBound(arrFileList)
            Call Sleep(50)
            strfilename = Trim(arrFileList(i))
            m_strFileName = Trim(strfilename)
            
            Call ImportExcelWithProgress(txtNotImported.Text & "\" & m_strFileName, False)

            If bIsFirstStepOK = True Then
                MoveFile strSourcePath & "\" & strfilename, strSucPath & "\" & strfilename
                logWriter.DeleteFile
            Else
                MoveFile strSourcePath & "\" & strfilename, strFailPath & "\" & strfilename
            End If
        Next
    Else
        MsgBox "Folder path must be entered", vbInformation, "Kingdee prompt"
        Exit Sub
    End If
End Sub

'bShowMsgBox signed that the sub will show a message box when exception found.
Private Sub ImportExcelWithProgress(fullfilename As String, bShowMsgBox As Boolean)
    m_WaitType = 1
    Set m_ProgBar = New Bar
    m_ProgBar.Show Me
    Call ImportToExcel(fullfilename, bShowMsgBox)
    logWriter.CloseFile
     m_ProgBar.Unload
    Set m_ProgBar = Nothing
End Sub
'将文件转移
Private Sub MoveFile(SourceFile As String, DestFile As String)
On Error GoTo EHandler

    Dim f As New FileSystemObject
    If f.FileExists(SourceFile) = True Then
        If f.FileExists(DestFile) = True Then
            f.DeleteFile DestFile, True
        End If
        
        SetAttr SourceFile, vbNormal
        f.MoveFile SourceFile, DestFile
    End If
    Set f = Nothing
    Exit Sub
EHandler:
    MsgBox "Move file failed:" & Err.Description
    Err.Clear
End Sub

Private Sub CmdSelectInvoice_Click()
    cmdlg.Filter = "csv text file|*.csv"
    cmdlg.ShowOpen
    
    txtInvoiceFilePath.Text = cmdlg.filename
End Sub

Private Sub cmdSelectVoucher_Click()
    cmdlg.Filter = "csv text file|*.csv"
    cmdlg.ShowOpen
    txtVoucherFilePath.Text = cmdlg.filename
End Sub

Private Sub cmdSelect_Click()
    Dim ssql As String
    Dim ssql1 As String
    Dim oconnect As Object
    Dim rs As ADODB.Recordset
    
    Dim i, j As Integer

    
    ssql = "select distinct t1.FName,t2.FDate,t2.FTransDate,t3.fAccountID,(ISNULL(t2.FReference,''))FReference,t4.Fname,(ISNULL(t3.FExplanation,''))FExplanation"
    ssql = ssql & vbCrLf & "from t_Voucher t2 inner join t_VoucherGroup t1 on t1.FGroupID=t2.FGroupID"
    ssql = ssql & vbCrLf & "inner join t_VoucherEntry t3 on t2.FVoucherID=t3.FVoucherID"
    ssql = ssql & vbCrLf & "inner join t_Currency t4 on t3.FCurrencyID=t4.FCurrencyID"
    ssql = ssql & vbCrLf & "where t3.FAccountID=20002"
    
    
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
     
    
    With fpSpread1
        .row = 1: .BackColor = &HC0C0C0
        .row = 2: .BackColor = &HC0C0C0
        .row = 5: .BackColor = &HC0C0C0
        .row = 6: .BackColor = &HC0C0C0
        .row = 3: .BackColor = &H80000001
        .row = 3: .ForeColor = &H80000005
        .row = 7: .BackColor = &H8000&
        .row = 7: .ForeColor = &H80000005

        .MaxCols = rs.Fields.Count + 3
        .MaxRows = 10
    
          For i = 1 To rs.RecordCount
            For j = 0 To rs.Fields.Count - 1
                .row = i + 3
                .Col = j + 2
                .Text = rs.Fields(j).Value
            Next
            rs.MoveNext
        Next


    End With
    
End Sub

Private Sub cmdExport_Click()

    Dim excelApp As New excel.Application
    Dim excelBook As New excel.Workbook
    Dim excelSheet As New excel.Worksheet
    Dim fpcol, fprow, i, j As Integer

     
    '将fpspread中的数据导出到excel中
    
    Set excelApp = CreateObject("excel.application")
    excelApp.Visible = False
    Set excelBook = excelApp.Workbooks.Add
    Set excelSheet = excelApp.ActiveSheet


    With fpSpread1
        fpcol = .MaxCols
        fprow = .MaxRows
   
    For i = 1 To fprow
        For j = 1 To fpcol
        fpSpread1.Col = j
        fpSpread1.row = i - 1
        excelSheet.Cells(i, j) = fpSpread1.Text
    Next j
    Next i
    
     End With
     
        CommonDialog1.filename = ""
        CommonDialog1.Filter = "csv |*.csv"
        CommonDialog1.ShowSave
        Text1.Text = CommonDialog1.filename
        ActiveWorkbook.SaveAs filename:=CommonDialog1.filename
        MsgBox "导出成功！"

    
End Sub


Private Sub Command1_Click()

    Dim ssql As String
    Dim oconnect As Object
    Dim rs As ADODB.Recordset
    Dim i, j As Integer
    
    Dim ssql1 As String
    Dim oconnect1 As Object
    Dim rs1 As ADODB.Recordset
'    Dim Text1 As Variant
    
       
'    If Text1.Text <> "" Then
        ssql1 = "select distinct t1.FName,t2.FDate,t2.FTransDate,(ISNULL(t2.FReference,''))FReference,"
        ssql1 = ssql1 & vbCrLf & "t4.Fname Currency, t2.fvoucherid"
        ssql1 = ssql1 & vbCrLf & "from t_Voucher t2 inner join t_VoucherGroup t1 on t1.FGroupID=t2.FGroupID"
        ssql1 = ssql1 & vbCrLf & "inner join t_VoucherEntry t3 on t2.FVoucherID=t3.FVoucherID"
        ssql1 = ssql1 & vbCrLf & "inner join t_Currency t4 on t3.FCurrencyID=t4.FCurrencyID"
        If Text1.Text <> "" Then
            ssql1 = ssql1 & vbCrLf & "where t2.FYear like '%" & Text1.Text & "%'"
        End If

        Set oconnect1 = CreateObject("K3Connection.AppConnection")
        Set rs1 = oconnect1.Execute(ssql1)


            With fpSpread2
                .MaxCols = rs1.Fields.Count
                .MaxRows = rs1.RecordCount
    
                For i = 1 To rs1.RecordCount
                    For j = 0 To rs1.Fields.Count - 1
                    .row = i
                    .Col = j + 2
                    .Text = rs1.Fields(j).Value
        
                    Next
                    rs1.MoveNext
                Next
            End With

'    Else
'        ssql = "select distinct t1.FName,t2.FDate,t2.FTransDate,(ISNULL(t2.FReference,''))FReference,"
'        ssql = ssql & vbCrLf & "t4.Fname Currency, t2.fvoucherid"
'        ssql = ssql & vbCrLf & "from t_Voucher t2 inner join t_VoucherGroup t1 on t1.FGroupID=t2.FGroupID"
'        ssql = ssql & vbCrLf & "inner join t_VoucherEntry t3 on t2.FVoucherID=t3.FVoucherID"
'        ssql = ssql & vbCrLf & "inner join t_Currency t4 on t3.FCurrencyID=t4.FCurrencyID"
'
'
'            Set oconnect = CreateObject("K3Connection.AppConnection")
'            Set rs = oconnect.Execute(ssql)
'
'
'            With fpSpread2
'                .MaxCols = rs.Fields.Count
'                .MaxRows = rs.RecordCount
'
'                  For i = 1 To rs.RecordCount
'                    For j = 0 To rs.Fields.Count - 1
'                        .row = i
'                        .Col = j + 2
'                        .Text = rs.Fields(j).Value
'
'                    Next
'                    rs.MoveNext
'                Next
'            End With

'   End If
  
End Sub

Private Sub Command2_Click()
    Dim ssql As String
    Dim oconnect As Object
    Dim rs As ADODB.Recordset
    Dim ssql1 As String
    Dim oconnect1 As Object
    Dim rs1 As ADODB.Recordset

    Dim excelApp As Object
    Dim excelBook As Object
    Dim excelSheet As Object
               

    Dim i, j, x, y As Integer
    
    Set shlShell = New Shell32.Shell
    '最后，显示对话框并返回结果：
    Set shlFolder = shlShell.BrowseForFolder(Me.hwnd, "Select a Folder", BIF_RETURNONLYFSDIRS)
                
                
    With fpSpread2
        For i = 1 To .MaxRows
             .row = i
             .Col = 1
             If .Value = 1 Then ' 选中凭证后
 
               
             '         打开excel模板
             Set excelApp = CreateObject("excel.application")
             excelApp.Visible = False
             
             Set excelBook = excelApp.Workbooks.Open("C:\Documents and Settings\Administrator\桌面\muban.xlt")
             Set excelSheet = excelApp.ActiveSheet
             
             .Col = 6
                    
             '         插入表头信息
             ssql = "select distinct t1.FName,t2.FDate,t2.FTransDate,'2480' as FCompanyCode,(ISNULL(t2.FReference,''))FReference,t4.Fname Currency,(ISNULL(t3.FExplanation,''))FExplanation"
             ssql = ssql & vbCrLf & "from t_Voucher t2 inner join t_VoucherGroup t1 on t1.FGroupID=t2.FGroupID"
             ssql = ssql & vbCrLf & "inner join t_VoucherEntry t3 on t2.FVoucherID=t3.FVoucherID and FEntryID=0"
             ssql = ssql & vbCrLf & "inner join t_Currency t4 on t3.FCurrencyID=t4.FCurrencyID"
             ssql = ssql & vbCrLf & "Where t2.FVoucherID = " & .Value
             
             Set oconnect = CreateObject("K3Connection.AppConnection")
             Set rs = oconnect.Execute(ssql)
                           
             excelSheet.Range("b4") = rs.Fields("FName")
             excelSheet.Range("c4") = rs.Fields("FDate")
             excelSheet.Range("d4") = rs.Fields("FTransDate")
             excelSheet.Range("e4") = rs.Fields("FCompanyCode")
             excelSheet.Range("f4") = rs.Fields("FReference")
             excelSheet.Range("g4") = rs.Fields("Currency")
             excelSheet.Range("h4") = rs.Fields("FExplanation")
             
             '         插入表体信息
             ssql1 = " select (case t2.fdc when 1 then 40 else 50 end)FDC,t5.FNumber FAccountCode,'2480' as FCompanyCode,(ISNULL(t4.FNumber,'')) FCostCenter,"
             ssql1 = ssql1 & vbCrLf & "(ISNULL(t4.FNumber,'')) FProfitCenter,(ISNULL(t4.FName,''))FBusinessArea,t2.FAmountFor,(ISNULL(t2.FExplanation,''))FExplanation"
             ssql1 = ssql1 & vbCrLf & "from t_Voucher t1 inner join t_VoucherEntry t2 on t1.FVoucherID=t2.FVoucherID"
             ssql1 = ssql1 & vbCrLf & "inner join t_ItemDetail t3 on t2.FDetailID=t3.FDetailID"
             ssql1 = ssql1 & vbCrLf & "inner join t_Account t5 on t2.FAccountID=t5.FAccountID"
             ssql1 = ssql1 & vbCrLf & "left join t_Item t4 on t3.F2030=t4.FItemID and t4.FItemClassID=2030"
             ssql1 = ssql1 & vbCrLf & "Where t1.FVoucherID = " & .Value
             
             Set oconnect1 = CreateObject("K3Connection.AppConnection")
             Set rs1 = oconnect1.Execute(ssql1)
             
                  
             Dim iCurrentRow As Integer
             iCurrentRow = 8
             With excelSheet
              For x = 1 To rs1.RecordCount
                 For y = 2 To 9
                    excelSheet.Cells(x + iCurrentRow - 1, y) = rs1.Fields(y - 2).Value
                 Next
                 rs1.MoveNext
              Next
              
             '       设置保存路径
'                CommonDialog2.filename = ""
'                CommonDialog2.Filter = "xls|*.xls"
'                CommonDialog2.ShowSave
             
             
             excelApp.ActiveWorkbook.SaveAs filename:=shlFolder.Self.path & "\" & Format(Now, "yyyymmddhhmmss") & ".xls"
              MsgBox "导出成功！"
             Set excelSheet = Nothing
             Set excelBook = Nothing
             excelApp.Quit
             Set excelApp = Nothing
        End With
      End If

      Next
      
             

   End With
         

End Sub

Private Sub Command3_Click()
 Dim ssql As String
    Dim oconnect As Object
    Dim rs As ADODB.Recordset
    Dim i, j As Integer
    

        ssql = "select distinct t1.FName,t2.FDate,t2.FTransDate,(ISNULL(t2.FReference,''))FReference,"
        ssql = ssql & vbCrLf & "t4.Fname Currency, t2.fvoucherid"
        ssql = ssql & vbCrLf & "from t_Voucher t2 inner join t_VoucherGroup t1 on t1.FGroupID=t2.FGroupID"
        ssql = ssql & vbCrLf & "inner join t_VoucherEntry t3 on t2.FVoucherID=t3.FVoucherID"
        ssql = ssql & vbCrLf & "inner join t_Currency t4 on t3.FCurrencyID=t4.FCurrencyID"
        ssql = ssql & vbCrLf & "where t2.FYear like '%" & Text1.Text & "%'"
    
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
    
    
    With fpSpread2
        .MaxCols = rs.Fields.Count
        .MaxRows = rs.RecordCount
    
          For i = 1 To rs.RecordCount
            For j = 0 To rs.Fields.Count - 1
                .row = i
                .Col = j + 2
                .Text = rs.Fields(j).Value
                                
            Next
            rs.MoveNext
        Next


    End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    If MMTS.CheckMts(1) = True Then
        lblAccountName.Caption = MMTS.AcctName
        dtImportedDate = FormatDateTime(Now, vbShortDate)
        bIsImported = False
    Else
        iExit = 1
        Unload Me
        Exit Sub
    End If
    
    txtNotImported.Text = App.path & "\Not_Imported"
    txtImported.Text = App.path & "\Imported"
    txtFailure.Text = App.path & "\Failure"
    
    fpSpread2.MaxCols = 0
    fpSpread2.MaxRows = 0
    
    
    '向系统托盘添加图标
    With NICONDATA
         .cbSize = Len(NICONDATA)
         .hwnd = frmMain.hwnd
         .uID = vbNull
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = frmMain.Icon
         .szTip = "kingdee automic import" & vbNullChar     ' vBNullChar : 用于单个 Null 字符的 Basic 常数 (ASCII value 0); 等效于 Chr$(0)
    End With
    Shell_NotifyIcon NIM_ADD, NICONDATA
    Call ReadFtpSettings
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bIsImporting = False Then
    Else
        Exit Sub
    End If
    
    '右下角图标左键单击控制窗体的显示
    If CLng(x / Screen.TwipsPerPixelX) = WM_LBUTTONDOWN Then
       '    X / Screen.TwipsPerPixelX : 多少象素点
       If frmMain.WindowState = 1 Or frmMain.WindowState = 0 Then
          frmMain.WindowState = 0
       Else
          frmMain.WindowState = 2
       End If
       
       If BMenuMainW = True Then
          frmMain.Hide
          BMenuMainW = Not BMenuMainW
       Else
          frmMain.Show
          BMenuMainW = Not BMenuMainW
       End If
       'MenuMainW.Checked = BMenuMainW
    End If
End Sub

Private Sub Form_Resize()
    If bIsImporting = False Then
    Else
        Exit Sub
    End If
    
    If frmMain.WindowState = 1 Then
        frmMain.Hide
        BMenuMainW = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If iExit = 0 Then
        If MsgBox("Are you sure you want to exit?", vbOKCancel + vbQuestion, "Kingdee Prompt") = vbOK Then
          tmrUpdate.Enabled = False
          Shell_NotifyIcon NIM_DELETE, NICONDATA     '删除图标
        Else
          Cancel = True
        End If
    End If
End Sub



Private Sub tmrUpdate_Timer()
    On Error GoTo EHandler
    
    If bIsImporting Then Exit Sub

    Dim dtPickTime As Date
    dtPickTime = FormatDateTime(DTPicker1.Value, vbShortTime)
    
    If Time < dtPickTime Then '小于定时器时间，直接退出
        Exit Sub
    End If
    
    If dtImportedDate < Date Then '第二天重置
      bIsImported = False
    End If

    If bIsImported = False Then
        If bIsImporting Then Exit Sub
        
        bIsImporting = True
        If txtFtpServerUrl.Text <> "" Then
            '执行Ftp文件下载
            DownloadFilesFromFtp txtNotImported.Text
        End If
        ImportFiles txtNotImported.Text, txtImported.Text, txtFailure.Text
    
        dtImportedDate = Date
        bIsImporting = False
        bIsImported = True
    End If
    Exit Sub
EHandler:
    bIsImporting = False
End Sub


'Copy from Manual Import Code
'bShowMsgBox signed that ,the function will show a message box when error raised
Private Function ImportToExcel(ByVal strfilename As String, Optional bShowMsgBox As Boolean = False) As Boolean
    Dim iRow As Long
    Dim strMsg As New StringBuilder
    Dim vctAllData As KFO.Vector 'Data list
    Dim Timestart As Date
    Dim DiffMinutes As Long
    Dim lngRowsCount As Long
    Dim objCreate As Object 'middle layer component object
    
    Dim progressSign As String 'for create error message
    Dim filenameWithoutPath As String
    filenameWithoutPath = GetFileNameWithoutPath(strfilename)
    
    progressSign = "Initial log file"
    logWriter.InitLogWithFileName txtFailure.Text, filenameWithoutPath
    
On Error GoTo HErr
    
    'ReadFile
    progressSign = "Read File " & strfilename
    Dim fileTypeFlag As String
    Dim vec As KFO.Vector
    'Set vec = ReadCSVFile(strfilename, fileTypeFlag)
    Set vec = ReadExcelFile(strfilename, fileTypeFlag)
    lngRowsCount = vec.Size
    logWriter.WriteLine "Parse file successful!"
    
    bIsFirstStepOK = True
    Timestart = Now

    Set vctAllData = New KFO.Vector
    m_ProgBar.SetBarMaxValue vec.Size
    m_ProgBar.ShowProgBar
    
    If vec.Size = 0 Then
        logWriter.WriteLine "Empty data"
        bIsFirstStepOK = False
        ImportToExcel = False
        Exit Function
    End If
    Dim dic As KFO.Dictionary
    If UCase(fileTypeFlag) = UCase("Journal Number") Then 'import k3 voucher
        logWriter.WriteLine "The file is voucher file!"
        progressSign = "Parse file"
        For iRow = 1 To vec.Size
           Set dic = vec(iRow)
           PackageVoucherData dic, iRow + 1, strMsg, vctAllData
           
           m_ProgBar.SetBarValue iRow
           DiffMinutes = DateDiff("s", Timestart, Now)
           m_ProgBar.SetMsg "please wait..." & vbCrLf & "Checking Row " & iRow & " of " & lngRowsCount & ",Total spent " & DiffMinutes & " seconds"
           DoEvents
        Next
        m_ProgBar.SetBarValueWithMax
        
        If strMsg.StringValue <> "" Then
            logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False

            ImportToExcel = False
            Exit Function
        End If
        
        progressSign = "transfer temporary data to temporary table"
        If InsertDataToTable(vctAllData, strMsg.StringValue) = False Then
            logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False
            ImportToExcel = False
            Exit Function
        End If
        logWriter.WriteLine "Insert voucher data to temp table!"
        progressSign = "create remoting object for save"
        Set objCreate = CreateObject("PF_Mid_ImportFunction.clsImportFunction")
        
        progressSign = "saving voucher"
        If objCreate.CreateVouchers(MMTS.PropsString, strUUID, strMsg) = False Then
            Call ExecSql("delete from PF_t_VoucherData where FUUID = '" & strUUID & "'")
            logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False
            ImportToExcel = False
            Exit Function
        End If
        logWriter.WriteLine "Save procedure is completed!"
    ElseIf UCase(fileTypeFlag) = UCase("Invoice Number") Then
        progressSign = "Parse file"
        For iRow = 1 To vec.Size
           Set dic = vec(iRow)
           PackageInvoice dic, iRow, strMsg, vctAllData
           m_ProgBar.SetBarValue iRow
           DiffMinutes = DateDiff("s", Timestart, Now)
           m_ProgBar.SetMsg "please wait..." & vbCrLf & "Checking Row " & iRow & " of " & lngRowsCount & ",Total spent " & DiffMinutes & " seconds"
           DoEvents
        Next
        m_ProgBar.SetBarValueWithMax
        
        If strMsg.StringValue <> "" Then
           logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False
            ImportToExcel = False
            Exit Function
        End If

        progressSign = "transfer temporary data to temporary table"
        If InsertInvoiceToTable(vctAllData, strMsg) = False Then
            logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False
            ImportToExcel = False
            Exit Function
        End If
        progressSign = "create empty invoice data package"
        Dim oDataSrv As Object
        Set oDataSrv = CreateObject("K3ClassTpl.DataSrv")
        oDataSrv.propstring = MMTS.PropsString
        oDataSrv.ClassTypeID = 1000002
        Dim vctBills As New KFO.Vector
        Dim bill As KFO.Dictionary
        Dim invNumber As String
        invNumber = ""
        Dim row As Long
        For row = 1 To vctAllData.Size
            If invNumber <> vctAllData(row)("FInvNumber") Then
                Set bill = oDataSrv.GetEmptyBill()
                vctBills.Add bill
                invNumber = vctAllData(row)("FInvNumber")
            End If
        Next
        progressSign = "create remoting object for save"
        Set objCreate = CreateObject("PF_Mid_ImportFunction.clsImportFunction")
        progressSign = "saving invoice"
        If objCreate.CreateInvoice(MMTS.PropsString, strUUID, vctBills, strMsg) = False Then
            ExecSql "Delete From PF_t_InvoiceData Where FUUID='" & strUUID & "'"
            logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False
            ImportToExcel = False
            Exit Function
        End If
    Else
        ImportToExcel = False
        logWriter.WriteLine "The file is neither a voucher file nor a invoice file!"
        Exit Function
    End If
    logWriter.WriteLine "The file has been imported!"
    ImportToExcel = True
    Exit Function
HErr:
    ImportToExcel = False
    bIsFirstStepOK = False

    logWriter.WriteLine "Failed on " & progressSign & ";" & Err.Description
    If bShowMsgBox = True Then
        MsgBox "Failed on " & progressSign & ":" & Err.Description, vbInformation + vbOKOnly, "Kingdee Prompt"
    End If
End Function

'导入凭证
Private Function ImportVouchers(ByVal strfilename As String, Optional bShowMsgBox As Boolean = False)
    Dim iRow As Long
    Dim iColumn As Long
    Dim iline As Long
    Dim strMsg As New StringBuilder
    Dim vctAllData As KFO.Vector
    Dim Timestart As Date
    Dim TimeEnd As Date
    Dim lngStartTime As Long
    Dim DiffMinutes As Long
    Dim lngRowsCount As Long
    Dim strSheetName As String
    Dim objCreate As Object
    Dim blnTemp As Boolean
    Dim progressSign As String 'for create error message
    
    Dim filenameWithoutPath As String
    filenameWithoutPath = GetFileNameWithoutPath(strfilename)
    
    progressSign = "Initial log file"
    logWriter.InitLogWithFileName txtFailure.Text, filenameWithoutPath
    
On Error GoTo HErr
    
    'ReadFile
    progressSign = "Read File " & strfilename
    Dim fileTypeFlag As String
    Dim vec As KFO.Vector
    'Set vec = ReadCSVFile(strfilename, fileTypeFlag)
    Set vec = ReadExcelFile(strfilename, fileTypeFlag)
    logWriter.WriteLine "Parse file successful!"
    
    bIsFirstStepOK = True
    blnTemp = True
    Timestart = Now
    lngStartTime = GetTickCount

    Set vctAllData = New KFO.Vector
    m_ProgBar.SetBarMaxValue vec.Size
    m_ProgBar.ShowProgBar
    
    If vec.Size = 0 Then
        blnTemp = False
        logWriter.WriteLine "Empty data"
        bIsFirstStepOK = False
        ImportVouchers = False
        Exit Function
    End If
    Dim dic As KFO.Dictionary
    If UCase(fileTypeFlag) = UCase("Journal Number") Then 'import k3 voucher
        logWriter.WriteLine "The file is voucher file!"
        progressSign = "Parse file"
        For iRow = 1 To vec.Size
           Set dic = vec(iRow)
           PackageVoucherData dic, iRow + 1, strMsg, vctAllData
           m_ProgBar.SetBarValue iRow
           TimeEnd = Now
           DiffMinutes = DateDiff("s", Timestart, TimeEnd)
           m_ProgBar.SetMsg "please wait..." & vbCrLf & "Checking Row " & iRow & " of " & lngRowsCount & ",Total spent " & DiffMinutes & " seconds"
           DoEvents
        Next
        m_ProgBar.SetBarValueWithMax
        
        If strMsg.StringValue <> "" Then
            blnTemp = False
            logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False

            ImportVouchers = False
            Exit Function
        End If
        
        progressSign = "transfer temporary data to temporary table"
        If InsertDataToTable(vctAllData, strMsg.StringValue) = True Then
            blnTemp = True
        Else
            blnTemp = False
            logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False

            ImportVouchers = False
            Exit Function
        End If
        logWriter.WriteLine "Insert voucher data to temp table!"
        progressSign = "create remoting object for save"
        Set objCreate = CreateObject("PF_Mid_ImportFunction.clsImportFunction")
        
        progressSign = "saving voucher"
        If objCreate.CreateVouchers(MMTS.PropsString, strUUID, strMsg) = True Then
            blnTemp = True
        Else
            blnTemp = False
            Call ExecSql("delete from PF_t_VoucherData where FUUID = '" & strUUID & "'")
            logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False

            ImportVouchers = False
            Exit Function
        End If
        logWriter.WriteLine "Save procedure is completed!"
    Else
        ImportVouchers = False
        logWriter.WriteLine "The file is not a voucher file!"
    End If
    logWriter.WriteLine "The file has been imported!"
    ImportVouchers = blnTemp
    Exit Function
HErr:
    blnTemp = False
    ImportVouchers = blnTemp
    bIsFirstStepOK = False

    logWriter.WriteLine "Failed on " & progressSign & ";" & Err.Description
    If bShowMsgBox = True Then
        MsgBox "Failed on " & progressSign & ":" & Err.Description, vbInformation + vbOKOnly, "Kingdee Prompt"
    End If
End Function


'导入发票
Private Function ImportInvoices(ByVal strfilename As String, Optional bShowMsgBox As Boolean = False)
    Dim iRow As Long
    Dim iColumn As Long
    Dim iline As Long
    Dim strMsg As New StringBuilder
    Dim vctAllData As KFO.Vector
    Dim Timestart As Date
    Dim TimeEnd As Date
    Dim lngStartTime As Long
    Dim DiffMinutes As Long
    Dim lngRowsCount As Long
    Dim strSheetName As String
    Dim objCreate As Object
    Dim blnTemp As Boolean
    Dim progressSign As String 'for create error message
    
    
    Dim filenameWithoutPath As String
    filenameWithoutPath = GetFileNameWithoutPath(strfilename)
    
    progressSign = "Initial log file"
    logWriter.InitLogWithFileName txtFailure.Text, filenameWithoutPath
    
On Error GoTo HErr
    
    'ReadFile
    progressSign = "Read File " & strfilename
    Dim fileTypeFlag As String
    Dim vec As KFO.Vector
    'Set vec = ReadCSVFile(strfilename, fileTypeFlag)
    Set vec = ReadExcelFile(strfilename, fileTypeFlag)
    logWriter.WriteLine "Parse file successful!"
    
    bIsFirstStepOK = True
    blnTemp = True
    Timestart = Now
    lngStartTime = GetTickCount

'    strMsg.StringValue = ""
    Set vctAllData = New KFO.Vector
    m_ProgBar.SetBarMaxValue vec.Size
    m_ProgBar.ShowProgBar
    
    If vec.Size = 0 Then
        blnTemp = False
        logWriter.WriteLine "Empty data"
        bIsFirstStepOK = False
        ImportInvoices = False
        Exit Function
    End If
    Dim dic As KFO.Dictionary
    If UCase(fileTypeFlag) = UCase("Invoice Number") Then
        progressSign = "Parse file"
        For iRow = 1 To vec.Size
           Set dic = vec(iRow)
           PackageInvoice dic, iRow, strMsg, vctAllData
           m_ProgBar.SetBarValue iRow
           TimeEnd = Now
           DiffMinutes = DateDiff("s", Timestart, TimeEnd)
           m_ProgBar.SetMsg "please wait..." & vbCrLf & "Checking Row " & iRow & " of " & lngRowsCount & ",Total spent " & DiffMinutes & " seconds"
           DoEvents
        Next
        m_ProgBar.SetBarValueWithMax
        
        If strMsg.StringValue <> "" Then
            blnTemp = False
           logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False

            ImportInvoices = False
            Exit Function
        End If

        progressSign = "transfer temporary data to temporary table"
        If InsertInvoiceToTable(vctAllData, strMsg) = True Then
            'Text1.Text = strFileName
            blnTemp = True
        Else
            blnTemp = False
            logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False

            ImportInvoices = False
            Exit Function
        End If
        progressSign = "create empty invoice data package"
        Dim oDataSrv As Object
        Set oDataSrv = CreateObject("K3ClassTpl.DataSrv")
        oDataSrv.propstring = MMTS.PropsString
        oDataSrv.ClassTypeID = 1000002
        Dim vctBills As New KFO.Vector
        Dim bill As KFO.Dictionary
        Dim invNumber As String
        invNumber = ""
        Dim row As Long
        For row = 1 To vctAllData.Size
            If invNumber <> vctAllData(row)("FInvNumber") Then
                Set bill = oDataSrv.GetEmptyBill()
                vctBills.Add bill
                invNumber = vctAllData(row)("FInvNumber")
            End If
        Next
        progressSign = "create remoting object for save"
        Set objCreate = CreateObject("PF_Mid_ImportFunction.clsImportFunction")
        progressSign = "saving invoice"
        If objCreate.CreateInvoice(MMTS.PropsString, strUUID, vctBills, strMsg) = True Then
            blnTemp = True
        Else
            blnTemp = False
            ExecSql "Delete From PF_t_InvoiceData Where FUUID='" & strUUID & "'"
            logWriter.WriteLine strMsg.StringValue
            bIsFirstStepOK = False

            ImportInvoices = False
            Exit Function
        End If
    Else
        ImportInvoices = False
        logWriter.WriteLine "The file is not a invoice file!"
    End If
    logWriter.WriteLine "The file has been imported!"
    ImportInvoices = blnTemp
    Exit Function
HErr:
    blnTemp = False
    ImportInvoices = blnTemp
    bIsFirstStepOK = False

    logWriter.WriteLine "Failed on " & progressSign & ";" & Err.Description
    If bShowMsgBox = True Then
        MsgBox "Failed on " & progressSign & ":" & Err.Description, vbInformation + vbOKOnly, "Kingdee Prompt"
    End If
End Function


Private Function ReadExcelFile(filename As String, ByRef fileTypeFlag As String) As KFO.Vector
    Dim iRow As Long
    Dim iColumn As Long
    Dim iline As Long
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    
    Dim Timestart As Date
    Dim TimeEnd As Date
    Dim lngStartTime As Long
    Dim DiffMinutes As Long
    Dim lngRowsCount As Long
    Dim lngColsCount As Long
    Dim strSheetName As String
    Dim objCreate As Object
    Dim blnTemp As Boolean
    
    Timestart = Now
On Error GoTo HErr
    Set xlApp = CreateObject("Excel.Application") '创建EXCEL对象
    Set xlBook = xlApp.Workbooks().Open(filename)
    Set xlSheet = xlBook.Worksheets(1) '打开EXCEL工作表
    xlApp.Visible = False '设置EXCEL对象可见（或不可见）
    
    Dim vec As New KFO.Vector
    Dim dic As KFO.Dictionary
    
'    strMsg.StringValue = ""
    lngRowsCount = xlSheet.UsedRange.Rows.Count
    lngColsCount = xlSheet.UsedRange.Columns.Count
    
    m_ProgBar.SetBarMaxValue lngRowsCount
    m_ProgBar.ShowProgBar
    
    fileTypeFlag = UCase(Trim(xlSheet.Cells(1, 1)))
    
    Dim iCol As Long
    For iRow = 2 To lngRowsCount
       If Trim(xlSheet.Cells(iRow, 1)) = "" Then
           Exit For
       End If
       
       Set dic = New KFO.Dictionary
       For iCol = 1 To lngColsCount
        dic(iCol) = Trim(xlSheet.Cells(iRow, iCol))
       Next iCol
       vec.Add dic
       m_ProgBar.SetBarValue iRow
       TimeEnd = Now
       DiffMinutes = DateDiff("s", Timestart, TimeEnd)
       'DiffMinutes = GetTickCount - lngStartTime
       m_ProgBar.SetMsg "please wait..." & vbCrLf & "Checking Row " & iRow & " of " & lngRowsCount & ",Total spent " & DiffMinutes & " seconds"
       DoEvents
    Next iRow
    
    m_ProgBar.SetBarValueWithMax
        
    xlBook.Close False
    xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
        
    Set ReadExcelFile = vec
    Exit Function
HErr:
    blnTemp = False
    Set ReadExcelFile = vec
    xlBook.Close False
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    MsgBox "Error raise when importing,please check the format of " & Mid(filename, 5, Len(filename)) & Err.Description, vbInformation + vbOKOnly, "Kingdee Prompt"
End Function

'Read a csv file and convert to a KFO Vector constructure
'flagText is file type flag
Private Function ReadCSVFile(filename As String, ByRef flagText As String) As KFO.Vector
On Error GoTo EHandler
    Dim vec As New KFO.Vector
    Dim dic As KFO.Dictionary
    Dim s() As String
    Dim textline As String

    Dim filenumber As Integer
    filenumber = FreeFile
    
    Open filename For Input As #filenumber
    Line Input #filenumber, textline
    flagText = SplitToKFODic(textline)(1)
    Do While Not EOF(filenumber)
        Line Input #filenumber, textline
        vec.Add SplitToKFODic(textline)
    Loop
    Close #filenumber
    Set ReadCSVFile = vec
    Exit Function
EHandler:
    Set ReadCSVFile = Nothing
    Close #filenumber
End Function


Private Function SplitToKFODic(txt As String) As KFO.Dictionary
    Dim s() As String
    s = Split(txt, ",")
    Dim dic As New KFO.Dictionary
    Dim i As Long
    For i = LBound(s) To UBound(s)
        dic(i + 1) = Trim(s(i))
    Next
    Set SplitToKFODic = dic
End Function

'为导入凭证向临时表PF_t_VoucherData写入要导入的凭证数据
Private Function InsertDataToTable(ByVal vctAllData As KFO.Vector, ByRef strMsg As String) As Boolean
    Dim i As Long
    Dim ssql As String
    Dim oconnect As Object
    Dim rs As ADODB.Recordset

   
    Dim strItem As String
    
    Dim objTypeLib As Object
    
    Dim dctCheck As KFO.Dictionary
    Dim dctTemp As KFO.Dictionary
    Dim vctTemp As KFO.Vector
    Dim strAllSQL As New StringBuilder
    
    '使用GUID作为一次事务的标识
    Set objTypeLib = CreateObject("Scriptlet.TypeLib")
    strUUID = CStr(objTypeLib.Guid)
    strUUID = Mid(strUUID, 1, InStr(1, strUUID, "}"))
    
    Set objTypeLib = Nothing
On Error GoTo HErr
    InsertDataToTable = False
    ssql = ""
    
    Set vctTemp = New KFO.Vector
    
    For i = vctAllData.LBound To vctAllData.UBound
         Set dctCheck = vctAllData(i)
    
         Set dctTemp = New KFO.Dictionary
         ssql = "insert  PF_t_VoucherData "
         ssql = ssql & vbCrLf & "values('" & vctAllData(i)("FVchNumber") & "',"
         ssql = ssql & "'" & vctAllData(i)("FYear") & "',"
         ssql = ssql & "'" & vctAllData(i)("FPeriod") & "',"
         ssql = ssql & "'" & vctAllData(i)("FDate") & "',"
         ssql = ssql & "'" & vctAllData(i)("FTransactionDate") & "',"
         ssql = ssql & "'" & vctAllData(i)("FVoucherCategory") & "',"
         ssql = ssql & "'" & vctAllData(i)("FLineNumber") & "',"
         ssql = ssql & "'" & vctAllData(i)("FCurrency") & "',"
         ssql = ssql & "'" & vctAllData(i)("FAccountNumber") & "',"
         ssql = ssql & "'" & vctAllData(i)("FDebitAmount") & "',"
         ssql = ssql & "'" & vctAllData(i)("FCreditAmount") & "',"
         ssql = ssql & "'" & vctAllData(i)("FAICustomer") & "',"
         ssql = ssql & "'" & vctAllData(i)("FAIMaterial") & "',"
         ssql = ssql & "'" & vctAllData(i)("FDescription") & "',"
         ssql = ssql & "'" & vctAllData(i)("FReference") & "',"
         ssql = ssql & "'" & strUUID & "')"
         
         dctTemp("sql") = ssql
         vctTemp.Add dctTemp
         Set dctTemp = Nothing
        
    Next
    
    If strMsg <> "" Then
       GoTo HErr
    Else
        'CmdImport.Enabled = True
    End If
    
    strAllSQL.Append "set nocount on"
    
    For i = vctTemp.LBound To vctTemp.UBound
        strAllSQL.Append vbCrLf & vctTemp(i)("sql")
        If i Mod 50 = 0 Then
       '    Debug.Print strAllSQL
           Set oconnect = CreateObject("K3Connection.AppConnection")
            oconnect.Execute (strAllSQL.StringValue)
            Set oconnect = Nothing
            strAllSQL.Remove 1, Len(strAllSQL.StringValue)
            strAllSQL.Append "set nocount on"
        End If
    Next
    
    If strAllSQL.StringValue <> "set nocount on" Then
      ' Debug.Print strAllSQL
        Set oconnect = CreateObject("K3Connection.AppConnection")
        oconnect.Execute (strAllSQL.StringValue)
        Set oconnect = Nothing
    End If
    InsertDataToTable = True
    Exit Function
HErr:
    InsertDataToTable = False
    ssql = "delete from PF_t_VoucherData where FUUID = '" & strUUID & "'"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    oconnect.Execute (ssql)
    Set oconnect = Nothing
    If strMsg <> "" Then
       strMsg = "Following Row has be import:" & vbCrLf & strMsg
    End If
    strMsg = strMsg & vbCrLf & CNulls(Err.Description, "")
End Function

'为导入发票准备临时表数据
Private Function InsertInvoiceToTable(ByVal vctAllData As KFO.Vector, ByRef strMsgTemp As StringBuilder) As Boolean
    Dim i As Long
    Dim ssql As String
    Dim oconnect As Object
    Dim rs As ADODB.Recordset
   
    Dim strItem As String
    Dim strMsg As String
    
    Dim objTypeLib As Object
    
    Dim dctCheck As KFO.Dictionary
    Dim dctTemp As KFO.Dictionary
    Dim vctTemp As KFO.Vector
    Dim strAllSQL As New StringBuilder
    
    '使用GUID作为一次导入事务的标识
    Set objTypeLib = CreateObject("Scriptlet.TypeLib")
    strUUID = CStr(objTypeLib.Guid)
    strUUID = Mid(strUUID, 1, InStr(1, strUUID, "}"))
    
    Set objTypeLib = Nothing
On Error GoTo HErr
    InsertInvoiceToTable = False
    ssql = ""
    
    Set vctTemp = New KFO.Vector
    
    For i = vctAllData.LBound To vctAllData.UBound
         Set dctCheck = vctAllData(i)
    
         Set dctTemp = New KFO.Dictionary
         ssql = "insert  PF_t_InvoiceData "
         ssql = ssql & vbCrLf & "values('" & vctAllData(i)("FInvNumber") & "',"
         ssql = ssql & "'" & vctAllData(i)("FInvDate") & "',"
         ssql = ssql & "'" & vctAllData(i)("FAccountingDate") & "',"
         ssql = ssql & "'" & vctAllData(i)("FReceiveDate") & "',"
         ssql = ssql & "'" & vctAllData(i)("FCustomerNumber") & "',"
         ssql = ssql & "'" & vctAllData(i)("FCustomerName") & "',"
         ssql = ssql & "'" & vctAllData(i)("FTaxRegistration") & "',"
         ssql = ssql & "'" & vctAllData(i)("FContactPerson") & "',"
         ssql = ssql & "'" & vctAllData(i)("fTelNumber") & "',"
         ssql = ssql & "'" & vctAllData(i)("FAddress1") & "',"
         ssql = ssql & "'" & vctAllData(i)("FAddress2") & "',"
         ssql = ssql & "'" & vctAllData(i)("FAddress3") & "',"
         ssql = ssql & "'" & vctAllData(i)("FAddress4") & "',"
         ssql = ssql & "'" & vctAllData(i)("FPostcode") & "',"
         ssql = ssql & "'" & vctAllData(i)("FMailAddress") & "',"
         ssql = ssql & "'" & vctAllData(i)("FBank") & "',"
         ssql = ssql & "'" & vctAllData(i)("FBankAccount") & "',"
         ssql = ssql & "'" & vctAllData(i)("FCurrency") & "',"
         ssql = ssql & "'" & vctAllData(i)("FExchangeRate") & "',"
         ssql = ssql & "'" & vctAllData(i)("FLineNumber") & "',"
         ssql = ssql & "'" & vctAllData(i)("FProductNumber") & "',"
         ssql = ssql & "'" & vctAllData(i)("FProductName") & "',"
         ssql = ssql & "'" & vctAllData(i)("FUoM") & "',"
         ssql = ssql & "'" & vctAllData(i)("FQty") & "',"
         ssql = ssql & "'" & vctAllData(i)("FPrice") & "',"
         ssql = ssql & "'" & vctAllData(i)("FDiscountRate") & "',"
         ssql = ssql & "'" & vctAllData(i)("FDiscountAmount") & "',"
         ssql = ssql & "'" & vctAllData(i)("FTaxRate") & "',"
         ssql = ssql & "'" & vctAllData(i)("FGoodsValue") & "',"
         ssql = ssql & "'" & vctAllData(i)("FTaxAmount") & "',"
         ssql = ssql & "'" & vctAllData(i)("FTotalAmount") & "',"
         ssql = ssql & "'" & vctAllData(i)("FRemark1") & "',"
         ssql = ssql & "'" & vctAllData(i)("FRemark2") & "',"
         ssql = ssql & "'" & vctAllData(i)("FRemark3") & "',"
         ssql = ssql & "'" & vctAllData(i)("FRemark4") & "',"
         ssql = ssql & "'" & strUUID & "')"
         
         dctTemp("sql") = ssql
         vctTemp.Add dctTemp
         Set dctTemp = Nothing
        
    Next
    
    strAllSQL.Append "set nocount on"
    
    For i = vctTemp.LBound To vctTemp.UBound
        strAllSQL.Append vbCrLf & vctTemp(i)("sql")
        If i Mod 50 = 0 Then
       '    Debug.Print strAllSQL
           Set oconnect = CreateObject("K3Connection.AppConnection")
            oconnect.Execute (strAllSQL.StringValue)
            Set oconnect = Nothing
            strAllSQL.Remove 1, Len(strAllSQL.StringValue)
            strAllSQL.Append "set nocount on"
        End If
    Next
    
    If strAllSQL.StringValue <> "set nocount on" Then
      ' Debug.Print strAllSQL
        Set oconnect = CreateObject("K3Connection.AppConnection")
        oconnect.Execute (strAllSQL.StringValue)
        Set oconnect = Nothing
    End If
    InsertInvoiceToTable = True
    Exit Function
HErr:
    InsertInvoiceToTable = False
    ssql = "delete from PF_t_InvoiceData where FUUID = '" & strUUID & "'"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    oconnect.Execute (ssql)
    Set oconnect = Nothing
    strMsg = strMsg & vbCrLf & CNulls(Err.Description, "")
    strMsgTemp.Append strMsg
End Function

'检查凭证是否存在
'strVchCate是凭证字
Private Function CheckVoucherNumber(ByVal strNumber As String, ByVal strVchCate As String, ByVal strYear As String, ByVal strPeriod As String) As Boolean
Dim ssql As String
Dim oconnect As Object
Dim rs As ADODB.Recordset
    ssql = "select t1.FNumber, t2.FName, t1.FYear, t1.FPeriod from t_Voucher t1"
    ssql = ssql & vbCrLf & "inner join t_VoucherGroup t2 on t1.FGroupID=t2.FGroupID"
    ssql = ssql & vbCrLf & "where t1.FNumber='" & strNumber & "' and t2.FName='" & strVchCate & "' and t1.FYear='" & strYear & "' and t1.FPeriod='" & strPeriod & "'"
    ssql = ssql & vbCrLf & "group by t1.FNumber, t2.FName, FYear, FPeriod"

    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
    If rs.EOF Then
       CheckVoucherNumber = False
    Else
       CheckVoucherNumber = True
    End If
    Set rs = Nothing
    Set oconnect = Nothing
End Function


'包装第iRow行的数据，并添加到vctAllData中
Private Sub PackageVoucherData(dic As KFO.Dictionary, ByVal iRow As Long, ByRef strMsgAll As StringBuilder, ByRef vctAllData As KFO.Vector)
    
    Dim tmpDict As KFO.Dictionary
    
    Dim strMsg As String
    Dim strTemp As String
    
    Dim rs As New Recordset
    Dim oConn As Object
    Set oConn = CreateObject("K3Connection.AppConnection")
    
    OrigAmount = 0
    
    Set tmpDict = New KFO.Dictionary
    
     'voucher number
    If Len(dic(1)) = 0 Or CheckVoucherNumber(dic(1), dic(6), dic(2), dic(3)) Then
        strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Journal number already exists in K/3" & vbCrLf
        tmpDict("FVchNumber") = ""
    Else
        tmpDict("FVchNumber") = dic(1)
    End If
    
    If Len(dic(2)) = 0 Or IsNumeric(dic(2)) = False Then
        strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Field:Year: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FYear") = ""
    Else
        If CLng(dic(2)) <> Val(dic(2)) Then
            strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Field:Year: Field type and the field length is incorrect" & vbCrLf
            tmpDict("FYear") = ""
        Else
            tmpDict("FYear") = dic(2)
        End If
    End If
    
    If Len(dic(3)) = 0 Or IsNumeric(dic(3)) = False Then
        strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Field:Period: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FPeriod") = ""
    Else
        If CLng(dic(3)) <> Val(dic(3)) Then
            strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Field:Period: Field type and the field length is incorrect" & vbCrLf
            tmpDict("FPeriod") = ""
        Else
            tmpDict("FPeriod") = dic(3)
        End If
    End If
    
    'FDate
    If Len(dic(4)) = 0 Or IsDate(dic(4)) = False Then
        strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Field:Voucher Date: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FDate") = ""
    Else
        tmpDict("FDate") = Format(dic(4), "yyyy/mm/dd")
        '检查日期是否在总账系统当前可录入单据日期之前
        Dim startdate As Date
        startdate = GetCurrentGLStartDate()
        If tmpDict("FDate") < startdate Then
            strMsg = strMsg & "Invoice Date is invalided,It must in current period!" & vbCrLf
        End If
    End If
       
    'FTransactionDate
    If Len(dic(5)) = 0 Or IsDate(dic(5)) = False Then
        strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Field:Transaction Date: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FTransactionDate") = ""
    Else
        tmpDict("FTransactionDate") = Format(dic(5), "yyyy/mm/dd")
    End If
    
    'Voucher Category
    If Len(dic(6)) = 0 Or TryParseVoucherGroup(dic(6), strTemp) = False Then
        strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Voucher Category " & dic(6) & " does not exist in K/3 make the import error" & vbCrLf
        tmpDict("FVoucherCategory") = ""
    Else
        tmpDict("FVoucherCategory") = strTemp
    End If
    
    tmpDict("FLineNumber") = dic(7)
        
    'FCurrency
    If Len(dic(8)) = 0 Or TryParseCurrency(dic(8), strTemp) = False Then
        strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Currency " & dic(8) & " does not exist in K/3 make the import error" & vbCrLf
        tmpDict("FCurrency") = ""
    Else
        tmpDict("FCurrency") = strTemp
    End If
        
    'FAccount
    If Len(dic(9)) = 0 Or CheckAccount(dic(9), strTemp) = False Then
        strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Account " & dic(9) & " does not exist in K/3 make the import error" & vbCrLf
        tmpDict("FAccountNumber") = ""
    Else
        tmpDict("FAccountNumber") = strTemp
        '检查科目是否支持指定的币种
        Dim rsAcc As ADODB.Recordset
        Set rsAcc = ExecSql("Select FNumber,FCurrencyID From t_Account Where FAccountID=" & strTemp)
        If Not (rsAcc.Fields("FCurrencyID").Value = 0 Or rsAcc.Fields("FCurrencyID").Value = 1 Or rsAcc.Fields("FCurrencyID").Value = Val(tmpDict("FCurrency"))) Then
            strMsg = strMsg & "Line:" & iRow & " K3 Account Number " & rsAcc.Fields("FNumber").Value & " not support currency " & dic(8) & vbCrLf
        End If
        
        tmpDict("FAICustomer") = ""
        tmpDict("FAIMaterial") = ""
            
        If CheckAccountProject(strTemp, rs) = True And (Len(dic(12)) <> 0 Or Len(dic(14)) <> 0) Then
            'Analysis Item Number: Customer
            rs.Filter = "fitemclassid = 1 and FAccountID = " & strTemp
            If rs.RecordCount > 0 Then
                If Len(dic(12)) <> 0 And TryParseCustomer(dic(12), strTemp) Then
                    tmpDict("FAICustomer") = strTemp
                Else
                    strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Analysis Item Number: Customer " & dic(12) & " does not exist in K/3 make the import error" & vbCrLf
'                    tmpDict("FAICustomer") = ""
                End If
            End If
            'Analysis Item Number: Product
            rs.Filter = "fitemclassid = 4 and FAccountID = " & strTemp
            If rs.RecordCount > 0 Then
                If Len(dic(14)) <> 0 And CheckProduct(dic(14), strTemp) Then
                    tmpDict("FAIMaterial") = strTemp
                Else
                    strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Analysis Item Number: Material " & dic(14) & " does not exist in K/3 make the import error" & vbCrLf
'                    tmpDict("FAIMaterial") = ""
                End If
            End If
'        ElseIf CheckAccountProject(strTemp, rs) = True And (Len(Trim(xlSheet.Cells(iRow, 12))) <> 0 Or Len(Trim(xlSheet.Cells(iRow, 14))) <> 0) Then
'
'        Else
'            tmpDict("FAICustomer") = ""
'            tmpDict("FAIMaterial") = ""
        End If
        
    End If
    
   
    
    'Debit
    If Len(dic(10)) = 0 Or IsNumeric(dic(10)) = False Then
        strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Field:Debit Amount: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FDebitAmount") = ""
    Else
        tmpDict("FDebitAmount") = dic(10)
    End If
    'Credit
    If Len(dic(11)) = 0 Or IsNumeric(dic(11)) = False Then
        strMsg = strMsg & "Journal Number:" & dic(1) & " Line:" & iRow & " Field:Credit Amount: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FCreditAmount") = ""
    Else
        tmpDict("FCreditAmount") = dic(11)
    End If

    tmpDict("FDescription") = dic(16)
    tmpDict("FReference") = dic(17)
    
    vctAllData.Add tmpDict
    Set tmpDict = Nothing
    
    strMsgAll.Append strMsg
    
End Sub

'打包一行应收发票数据
Private Sub PackageInvoice(dic As KFO.Dictionary, ByVal iRow As Long, ByRef strMsgAll As StringBuilder, ByRef vctAllData As KFO.Vector)
    Dim tmpDict As KFO.Dictionary
        
    Dim lngStartTime As Long
    Dim strTimeMsg As String
    
    Dim s() As String
    Dim strTemp As String
    Dim strMsg As String
    
On Error GoTo Err

    Set tmpDict = New KFO.Dictionary
    
    If Len(dic(1)) = 0 Or CheckInvoiceNumber(dic(1)) Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Invoice number already exists in K/3" & vbCrLf
        tmpDict("FInvNumber") = ""
    Else
        tmpDict("FInvNumber") = dic(1)
    End If
    
    'FInvDate
    If Len(dic(2)) = 0 Or IsDate(dic(2)) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:Invoice Date: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FInvDate") = ""
    Else
        tmpDict("FInvDate") = Format(dic(2), "yyyy/mm/dd")
        'Check Invoice date is valid.
        Dim startdate As Date
        startdate = GetCurrentArStartDate()
        If tmpDict("FInvDate") < startdate Then
            strMsg = strMsg & "Invoice Date is invalided,It must in current period!" & vbCrLf
        End If
    End If
    
    'FAccountingDate
    If Len(dic(3)) = 0 Or (Not IsDate(dic(3))) Then
        strMsg = strMsg & "Invoice Number: " & dic(1) & " Line " & iRow & " Field:Accounting Date: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FAccountingDate") = ""
    Else
        tmpDict("FAccountingDate") = Format(dic(3), "yyyy/mm/dd")
    End If
    
    
    'FReceiveDate
    If Len(dic(3)) = 0 Or (Not IsDate(dic(4))) Then
        strMsg = strMsg & "Invoice Number: " & dic(1) & " Line " & iRow & " Field:Receivable Date: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FReceiveDate") = ""
    Else
        tmpDict("FReceiveDate") = Format(dic(4), "yyyy/mm/dd")
    End If
    
    If Len(dic(5)) = 0 Then
        strMsg = strMsg & "Invoice Number: " & dic(1) & " Line " & iRow & " Field:Customer Number: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FCustomerNumber") = ""
    Else
        tmpDict("FCustomerNumber") = dic(5)
    End If
    
    If Len(dic(6)) = 0 Then
        strMsg = strMsg & "Invoice Number: " & dic(1) & " Line " & iRow & " Field:Customer Name: Customer Name is not allowed to be blank " & vbCrLf
        tmpDict("FCustomerName") = ""
    Else
        tmpDict("FCustomerName") = dic(6)
    End If
    
    'FTaxRegistration
    If Len(dic(7)) > 21 Then
        strMsg = strMsg & "Invoice Number: " & dic(1) & " Line " & iRow & " Length of Tax Registration Field in AR invoice is over than 21 characters" & vbCrLf
        tmpDict("FTaxRegistration") = ""
    Else
        tmpDict("FTaxRegistration") = dic(7)
    End If
    
    tmpDict("FContactPerson") = dic(8)
    tmpDict("fTelNumber") = dic(9)
    tmpDict("FAddress1") = dic(10)
    tmpDict("FAddress2") = dic(11)
    tmpDict("FAddress3") = dic(12)
    tmpDict("FAddress4") = dic(13)
    tmpDict("FPostcode") = dic(14)
    tmpDict("FMailAddress") = dic(15)
    tmpDict("FBank") = dic(16)
    tmpDict("FBankAccount") = dic(17)
    
    If Len(dic(18)) = 0 Or TryParseCurrency(dic(18), strTemp) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Currency " & dic(18) & " does not exist in K/3 make the import error" & vbCrLf
        tmpDict("FCurrency") = ""
    Else
        tmpDict("FCurrency") = strTemp
    End If
    
    If Len(dic(19)) = 0 Or IsNumeric(dic(19)) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:Exchange Rate(%): Field type and the field length is incorrect" & vbCrLf
        tmpDict("FExchangeRate") = ""
    Else
        tmpDict("FExchangeRate") = dic(19)
    End If
    
    tmpDict("FLineNumber") = dic(20)
    
    If Len(dic(21)) = 0 Or Len(dic(21)) > 100 Then
       strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Product code must be less than 100 characters " & vbCrLf
       tmpDict("FProductNumber") = ""
    Else
        tmpDict("FProductNumber") = dic(21)
    End If
    
    If Len(dic(22)) = 0 Or Len(dic(22)) > 255 Then
       strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Product name must be less than 255 characters " & vbCrLf
       tmpDict("FProductName") = ""
    Else
        tmpDict("FProductName") = dic(22)
    End If
    
    If Len(dic(23)) = 0 Or TryParseUoM(dic(23), strTemp) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " UoM " & dic(23) & " does not exist in K/3 make the import error" & vbCrLf
        tmpDict("FUoM") = ""
    Else
        tmpDict("FUoM") = strTemp
    End If
    
    If Len(dic(24)) = 0 Or IsNumeric(dic(24)) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:QTY: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FQty") = ""
    Else
        If Val(dic(24)) = 0 Then
            strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:QTY: Quantity is 0" & vbCrLf
        Else
            tmpDict("FQty") = dic(24)
        End If
    End If
    
    If Len(dic(25)) = 0 Or IsNumeric(dic(25)) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:Unit Price: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FPrice") = ""
    Else
        tmpDict("FPrice") = dic(25)
    End If
    
    If Len(dic(26)) = 0 Or IsNumeric(dic(26)) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:Discount Rate(%): Field type and the field length is incorrect" & vbCrLf
        tmpDict("FDiscountRate") = ""
    Else
        tmpDict("FDiscountRate") = dic(26)
    End If
    If Len(dic(27)) = 0 Or IsNumeric(dic(27)) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:Discount Amount: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FDiscountAmount") = ""
    Else
        tmpDict("FDiscountAmount") = dic(27)
    End If
    
    If Len(dic(28)) = 0 Or IsNumeric(dic(28)) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:Tax Rate(%): Field type and the field length is incorrect" & vbCrLf
        tmpDict("FTaxRate") = ""
    Else
        tmpDict("FTaxRate") = dic(28)
    End If
    If Len(dic(30)) = 0 Or IsNumeric(dic(30)) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:Tax Amount: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FTaxAmount") = ""
    Else
        tmpDict("FTaxAmount") = dic(30)
    End If
    
    If Len(dic(29)) = 0 Or IsNumeric(dic(29)) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:Goods Value: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FGoodsValue") = ""
    Else
        tmpDict("FGoodsValue") = dic(29)
    End If
    If Len(dic(31)) = 0 Or IsNumeric(dic(31)) = False Then
        strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " Field:Total Amount: Field type and the field length is incorrect" & vbCrLf
        tmpDict("FTotalAmount") = ""
    Else
        tmpDict("FTotalAmount") = dic(31)
    End If
    
    tmpDict("FRemark1") = dic(32)
    tmpDict("FRemark2") = dic(33)
    tmpDict("FRemark3") = dic(34)
    tmpDict("FRemark4") = dic(35)

    vctAllData.Add tmpDict
    Set tmpDict = Nothing
    
    strMsgAll.Append strMsg
    
    Exit Sub
    
Err:
    strMsg = strMsg & "Invoice Number:" & dic(1) & " Line:" & iRow & " has illegal format values" & vbCrLf
    
End Sub


'检查发票在临时表中是否存在,如果存在，可能导入过，也可能没导过
Private Function CheckInvoiceNumber(ByVal strNumber As String) As Boolean
Dim ssql As String
Dim oconnect As Object
Dim rs As ADODB.Recordset
    'sSQL = "select 1 from PF_t_InvoiceData where FInvNumber = '" & strNumber & "'"
    ssql = "Select 1 From ICSale Where FBillNo='" & strNumber & "'"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
    If rs.EOF Then
       CheckInvoiceNumber = False
    Else
       CheckInvoiceNumber = True
    End If
    Set rs = Nothing
    Set oconnect = Nothing
End Function

'检查币种是否存在，并传回找到的币种的内码
Private Function TryParseCurrency(ByVal strNumber As String, ByRef lCurrencyId) As Boolean
Dim ssql As String
Dim oconnect As Object
Dim rs As ADODB.Recordset
    ssql = "select FCurrencyID from t_currency where FName = '" & strNumber & "'"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
    If rs.EOF Then
       TryParseCurrency = False
       lCurrencyId = "0"
    Else
       lCurrencyId = rs("FCurrencyID")
       TryParseCurrency = True
    End If
    Set rs = Nothing
    Set oconnect = Nothing
End Function

'检查科目映射关系是否存在,并返回对应的K3科目的内码
'strNumber是其他系统的科目代码
'strTemp是K3科目内码
Private Function CheckAccount(ByVal strNumber As String, ByRef strTemp) As Boolean
Dim ssql As String
Dim oconnect As Object
Dim rs As ADODB.Recordset
'    ssql = "select FAccountNumberInK3 as FAccountID from t_COAMappingEntry where FAccountNumberInOrbit = '" & strNumber & "'"
    ssql = "SELECT FAccountID  FROM t_Account Where  FHelperCode = '" & strNumber & "'"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
    If rs.EOF Then
       CheckAccount = False
       strTemp = "0"
    Else
       strTemp = rs("FAccountID")
       CheckAccount = True
    End If
    Set rs = Nothing
    Set oconnect = Nothing
End Function


'检查计量单位
Private Function TryParseUoM(ByVal strNumber As String, ByRef lMeasureID) As Boolean
Dim ssql As String
Dim oconnect As Object
Dim rs As ADODB.Recordset
    ssql = "select FMeasureUnitID from t_MeasureUnit where FNumber = '" & strNumber & "'"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
    If rs.EOF Then
       TryParseUoM = False
       lMeasureID = 0
    Else
       TryParseUoM = True
       lMeasureID = rs("FMeasureUnitID")
    End If
    Set rs = Nothing
    Set oconnect = Nothing
End Function

'********  Added by Nicky  2010-5-19  **********  begin
Private Function TryParseCustomer(ByVal strNumber As String, ByRef lngItemID As String) As Boolean
Dim ssql As String
Dim oconnect As Object
Dim rs As ADODB.Recordset
    ssql = "select FItemID from t_item where fitemclassid = 1 and fnumber = '" & strNumber & "' and fdetail = 1"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
    If rs.EOF Then
       TryParseCustomer = False
       lngItemID = 0
    Else
       TryParseCustomer = True
       lngItemID = rs("FItemID")
    End If
    Set rs = Nothing
    Set oconnect = Nothing
End Function
'********  Added by Nicky  2010-5-19  **********  end

'********  Added by Nicky  2010-5-19  **********  begin
Private Function CheckProduct(ByVal strNumber As String, ByRef lngDeptID As String) As Boolean
Dim ssql As String
Dim oconnect As Object
Dim rs As ADODB.Recordset
    ssql = "select FItemID from t_item where fitemclassid = 4 and fnumber = '" & strNumber & "' and fdetail = 1"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
    If rs.EOF Then
       CheckProduct = False
       lngDeptID = 0
    Else
       CheckProduct = True
       lngDeptID = rs("FItemID")
    End If
    Set rs = Nothing
    Set oconnect = Nothing
End Function
'********  Added by Nicky  2010-5-19  **********  end

Private Function TryParseVoucherGroup(ByVal strNumber As String, ByRef voucherGroupID) As Boolean
Dim ssql As String
Dim oconnect As Object
Dim rs As ADODB.Recordset
    ssql = "select FGroupID from t_VoucherGroup where FName = '" & strNumber & "'"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
    If rs.EOF Then
       TryParseVoucherGroup = False
    Else
       voucherGroupID = rs("FGroupID")
       TryParseVoucherGroup = True
    End If
    Set rs = Nothing
    Set oconnect = Nothing
End Function

'Private Function CheckVendorCode(ByVal strNumber As String, ByRef lngVendorID) As Boolean
'Dim sSQL As String
'Dim oConnect As Object
'Dim rs As ADODB.Recordset
'
'
'    sSQL = "select FVendorID from t_BurVendorMappingEntry where FPOSVendorCode = '" & strNumber & "'"
'
'    Set oConnect = CreateObject("K3Connection.AppConnection")
'    Set rs = oConnect.Execute(sSQL)
'    If rs.EOF Then
'       CheckVendorCode = False
'       lngVendorID = -1
'    Else
'       CheckVendorCode = True
'       lngVendorID = rs("FVendorID")
'    End If
'    Set rs = Nothing
'    Set oConnect = Nothing
'
'End Function


Private Function CreateVoucher(strFUUID As String) As Boolean
    
    Dim result As KFO.Vector
    Dim dicVoucher As KFO.Dictionary
    Dim dicTempEntry As KFO.Dictionary
    Dim vctTemp As KFO.Vector
    Dim dctitem As KFO.Dictionary
    Dim vctVoucherEntry As KFO.Vector
    Dim strSQL As String, Msg As String
    Dim rs As ADODB.Recordset
    Dim rsDetails As ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim LVoucherID As String
    Dim vctVoucher As KFO.Vector
    Dim saveObject As Object
    Dim douOrgAmount As Double
    Dim oconnect As Object
    Dim iRow As Integer

    Dim objsave As Object

On Err GoTo Err

    Set objsave = CreateObject("BillDataAccess.GetData")
    
    strSQL = " SELECT FDocumentNo,FDate,FTransactionType,FCurrencyID,FExchangeRate,FDCr,FAccountID,FOriginalAmount,FReportingAmount," & vbCrLf
    'strSQL = strSQL & "FCostCenterID,FVendorID,FDescription,T2.FGroupID" & vbCrLf
    strSQL = strSQL & "FCostCenterID,FCustomerID,FVendorID,FDescription,T2.FGroupID" & vbCrLf
    strSQL = strSQL & "FROM t_BurBerryVchImport T1  LEFT JOIN t_BurCOAMapping T2 ON T1.FtrantypeID=T2.FTRANTYPE" & vbCrLf 'WHERE FUUID=''"
    strSQL = strSQL & "where t1.FUUID='" & strFUUID & "'" & vbCrLf
    'strSQL = strSQL & "order by FDocumentNo"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(strSQL)
   
    If rs.RecordCount = 0 Then
        Msg = "It's invalid to import the voucher that all amount are Zero!"

        GoTo Err
    End If
    
    m_ProgBar.SetBarMaxValue Val(rs.RecordCount)
    m_ProgBar.ShowProgBar
       
    Dim Timestart
    Timestart = Now
    Set vctVoucher = New KFO.Vector
    Do While rs.EOF = False
        
        Set dicVoucher = New KFO.Dictionary
        LVoucherID = rs("FDocumentNo")
        dicVoucher("FDate") = rs("FDate")                                       'Voucher head
        dicVoucher("FGroupID") = rs("FGroupID")
        dicVoucher("FPreparerID") = lUserID
        dicVoucher("FReference") = rs("FDocumentNo")
        If rs("FTransactiontype") = "S1" Then
            dicVoucher("FInternalInd") = "ar"
        ElseIf rs("FTransactiontype") = "P1" Or rs("FTransactiontype") = "P2" Then
            dicVoucher("FInternalInd") = "ap"
        End If
        
        Set dicTempEntry = New KFO.Dictionary                                   'Voucher Entry
        Set vctVoucherEntry = New KFO.Vector
        
        For i = 1 To rs.RecordCount
            If rs.EOF = True Then
                GoTo NextVoucher
            End If
            If LVoucherID <> rs("FDocumentNo") And rs("FDocumentNo") <> "" Then   'if dr<>cr then
                If douOrgAmount <> 0 Then
                    Set dicTempEntry = New KFO.Dictionary
                    dicTempEntry("FExplanation") = rs("FDescription")
                    dicTempEntry("FAccountID") = rs("FAccountID")
                    dicTempEntry("FCurrencyID") = rs("FCurrencyID")
                    dicTempEntry("FDC") = 1
                    
                    dicTempEntry("FExchangeRate") = rs("FExchangeRate")
                    dicTempEntry("FAmountFor") = Format(douOrgAmount, "###0.00")
                    dicTempEntry("FAmount") = Format(douOrgAmount * rs("FExchangeRate"), "##0.00")
                    vctVoucherEntry.Add dicTempEntry
                End If
                GoTo NextVoucher
            End If
            
            Set dicTempEntry = New KFO.Dictionary
            dicTempEntry("FExplanation") = rs("FDescription")
            dicTempEntry("FAccountID") = rs("FAccountID")
            dicTempEntry("FCurrencyID") = rs("FCurrencyID")
            dicTempEntry("FExchangeRate") = rs("FExchangeRate")
            
            If rs("FDCr") = "Dr" Then
                dicTempEntry("FDC") = 1
                douOrgAmount = -rs("FOriginalAmount") + douOrgAmount
            Else
                dicTempEntry("FDC") = 0
                douOrgAmount = rs("FOriginalAmount") + douOrgAmount
            End If
            
            dicTempEntry("FAmountFor") = rs("FOriginalAmount")
            dicTempEntry("FAmount") = rs("FReportingAmount")
    
            Set vctTemp = New KFO.Vector
            Set dctitem = New KFO.Dictionary
            
            If CheckAccountProject(2, rs("FAccountID")) = True Then
                If rs("FCostCenterID") <> 0 Then                 'Voucher Item
                    Set dctitem = New KFO.Dictionary
                    dctitem("FItemClassID") = 2
                    dctitem("FItemID") = rs("FCostCenterID")
                    vctTemp.Add dctitem
                Else
                    Msg = Msg & "Row " & CStr(i + 1) & ": The value in Cost Center(Department) column is a invalid data!" & vbCrLf
                    'GoTo Err
                End If
            End If
            
            If CheckAccountProject(1, rs("FAccountID")) = True Then
                If rs("FCustomerID") <> 0 Then      'Added by Nicky 2010-5-20
                    Set dctitem = New KFO.Dictionary
                    dctitem("FItemClassID") = 1
                    dctitem("FItemID") = rs("FCustomerID")
                    vctTemp.Add dctitem
                Else
                    Msg = Msg & "Row " & CStr(i + 1) & ": The value in Cost Center(Customer) column is a invalid data!" & vbCrLf
                    'GoTo Err
                End If
            End If
            
            If CheckAccountProject(8, rs("FAccountID")) = True Then  'Added by Nicky 2010-5-28
                If rs("FVendorID") <> 0 Then
                    Set dctitem = New KFO.Dictionary
                    dctitem("FItemClassID") = 8
                    dctitem("FItemID") = rs("FVendorID")
                    vctTemp.Add dctitem
                Else
                    If UCase(rs("FTransactionType")) <> UCase("P1") And UCase(rs("FTransactionType")) <> UCase("P2") Then
                        Set dctitem = New KFO.Dictionary
                        dctitem("FItemClassID") = 8
                        strSQL = "select FItemID from t_item where fitemclassid=8 and fnumber='01.9000000021'"
                        Set oconnect = CreateObject("K3Connection.AppConnection")
                        Dim rsAdd As ADODB.Recordset
                        Set rsAdd = oconnect.Execute(strSQL)
                        dctitem("FItemID") = CLng(rsAdd("FItemID"))
                        vctTemp.Add dctitem
                    Else
                        Msg = Msg & "Row " & CStr(i + 1) & ": The value in Vendor column is a invalid data!" & vbCrLf
                    End If
                End If
            End If
'                If rs("FVendorID") <> 0 Then
'                    Set dctitem = New KFO.Dictionary
'                    dctitem("FItemClassID") = 8
'                    dctitem("FItemID") = rs("FVendorID")
'                    vctTemp.Add dctitem
''                Else
''                        Msg = Msg & "Row " & CStr(i + 1) & ": The value in Vendor column is a invalid data!" & vbCrLf
'
'                End If

            If vctTemp.UBound > 0 Then
                Set dicTempEntry("_Details") = vctTemp
            End If
            vctVoucherEntry.Add dicTempEntry
            
            j = j + 1
            m_ProgBar.SetBarValue j   '进度条
            Dim TimeEnd
            TimeEnd = Now
            Dim DiffMinutes
            DiffMinutes = DateDiff("s", Timestart, TimeEnd)
            m_ProgBar.SetMsg "please wait..." & vbCrLf & "Checking Row " & j & " of " & Val(rs.RecordCount) & ",Total spent " & DiffMinutes & "minutes"
            DoEvents
            
            
            rs.MoveNext
        Next
        
NextVoucher:
        Set dicVoucher("_Entries") = vctVoucherEntry
        douOrgAmount = 0
        vctVoucher.Add dicVoucher
    Loop
    
    If Len(Msg) <> 0 Then
        GoTo Err
    End If
        
'''--生成凭证
  '  Set saveObject = CreateObject("SerCreateVoucher.ClsCreateVoucher")
    
 Set saveObject = CreateObject("testVoucher.ClsTestVoucher")  'TEST 使用
    
    Set result = saveObject.Create(MMTS.PropsString, vctVoucher, strFUUID, Msg) 'Modify by rock -Add "StrFUUID"
    
    If result Is Nothing Then
        Dim ssql As String
        ssql = "delete from t_BurBerryVchImport where FUUID = '" & strUUID & "'"
        Set oconnect = CreateObject("K3Connection.AppConnection")
        oconnect.Execute (ssql)
        CreateVoucher = False
        logWriter.WriteLine Msg


    Else
        logWriter.WriteLine Msg
        CreateVoucher = True
    End If

    Set result = Nothing
    Set dicVoucher = Nothing
    Set dicTempEntry = Nothing
    Set vctTemp = Nothing
    Set dctitem = Nothing
    Set vctVoucherEntry = Nothing
    Set rs = Nothing
    Set rsDetails = Nothing
    Set vctVoucher = Nothing
    Set saveObject = Nothing
Exit Function
Err:
    ssql = "delete from t_BurBerryVchImport where FUUID = '" & strUUID & "'"
    Set oconnect = CreateObject("K3Connection.AppConnection")
    oconnect.Execute (ssql)
    CreateVoucher = False
    logWriter.WriteLine Msg

    'MsgBox Err.Description, vbOKOnly + vbQuestion, "Kingdee Prompt"
    Set result = Nothing
    Set dicVoucher = Nothing
    Set dicTempEntry = Nothing
    Set vctTemp = Nothing
    Set dctitem = Nothing
    Set vctVoucherEntry = Nothing
    Set rs = Nothing
    Set rsDetails = Nothing
    Set vctVoucher = Nothing
    Set saveObject = Nothing

End Function


''
'Private Function ItemCheck(ByVal dicAllData As KFO.Dictionary, ByRef strItem As String) As Boolean
'
'
'        If dicAllData("CurrencyID") = 0 Then
'            strItem = "Currency is not Exists!" & vbCrLf & strMsg
'            ItemCheck = True
'        End If
'        If dicAllData("AccountID") = 0 Then
'            strItem = "Account is not Exists!" & vbCrLf & strMsg
'            ItemCheck = True
'        End If
''        If dicAllData("CostCenterID") = 0 Then
''            strItem = "Cost Center is not Exists!" & vbCrLf & strMsg
''            ItemCheck = True
''        End If
''        If dicAllData("VendorID") = 0 Then
''            strItem = "Vendor is not Exists!" & vbCrLf & strMsg
''            ItemCheck = True
''        End If
''        If dicAllData("GroupID") = 0 Then
''            strItem = "Voucher Cate is not Exists!" & vbCrLf & strMsg
''            VoucherCheck = True
''        End If
'
'    Set dicAllData = Nothing
'
'End Function

'
Private Sub m_ProgBar_Active()
    Select Case m_WaitType
       Case 1
            Dim myWaitCur As WaitCur
            Set myWaitCur = New WaitCur
            m_ProgBar.SetMsg "please wait..."
'          If ImportToExcel(m_strFileName) = True Then
'             m_ProgBar.Unload
'          Else
'             m_ProgBar.Unload
'          End If
          
       Case 2

            Set myWaitCur = New WaitCur
            m_ProgBar.SetMsg "please wait..."
            
            If CreateVoucher(strUUID) = True Then
               m_ProgBar.Unload
            Else
               m_ProgBar.Unload
          End If
    End Select
End Sub

''写入日志
''strFileName是被导入的文件的文件名,这里作为输入参数用来构造成日志文件名的一部分
'Public Sub WriteLine(ByVal strLog As String, ByVal strfilename As String)
'    Dim txtFile As Object
'    Dim fso As New FileSystemObject
'    Dim filepath As String
'
'    Dim strDateTimeToString
'
'    strDateTimeToString = Format(CStr(Now), "yyyymmddhhmmss")
'
'    If fso.FolderExists(txtFailure.Text) = False Then
'        fso.CreateFolder (txtFailure.Text)
'    End If
'
'    filepath = txtFailure.Text & "\" & strfilename & " - error log - " & strDateTimeToString & ".txt"
'
'    If Dir(filepath) = "" Then '判断文件是否存在，不存在就创建，存在就不创建
'        Open filepath For Append As #1
'        Close #1
'    End If
'
'    Set txtFile = fso.OpenTextFile(filepath, 8, True, 0)
'
'    txtFile.WriteLine strLog
'
'    txtFile.Close
'    Set txtFile = Nothing
'    Set fso = Nothing
'
'   'Shell "notepad " & filePath, vbNormalFocus
'
'End Sub

''分别得到编码为9001及9002的核算项目类别的内码
'Private Sub GetItemClassID(ByRef lngPOSTranTypeID As Long, _
'                           ByRef lngPOSSubTypeID As Long)
''Private m_lngPOSTranTypeID As Long
''Private m_lngPOSSubTypeID As Long
'    Dim sSQL As String
'    Dim oConnect As Object
'    Dim rs As ADODB.Recordset
'
'    Dim lngTimeStart As Long
'    lngTimeStart = GetTickCount
'    sSQL = "select fitemclassid from t_itemclass where fnumber = '9001'"
'
'    Set oConnect = CreateObject("K3Connection.AppConnection")
'    Set rs = oConnect.Execute(sSQL)
'    If rs.EOF Then
'       lngPOSTranTypeID = 0
'    Else
'       lngPOSTranTypeID = rs("fitemclassid")
'    End If
'    Set rs = Nothing
'    Set oConnect = Nothing
'
'    sSQL = "select fitemclassid from t_itemclass where fnumber = '9002'"
'    Set oConnect = CreateObject("K3Connection.AppConnection")
'    Set rs = oConnect.Execute(sSQL)
'    If rs.EOF Then
'       lngPOSSubTypeID = 0
'    Else
'       lngPOSSubTypeID = rs("fitemclassid")
'    End If
'    Set rs = Nothing
'    Set oConnect = Nothing
'   ' MsgBox GetTickCount - lngTimeStart
'End Sub

'确定核算项目  Added by Nicky
'检查科目的核算项目是否不是客户及物料之外的项目核算
'此项限制暂时不楚来源.
Private Function CheckAccountProject(ByVal lAccount As String, ByRef rs As Recordset) As Boolean
    Dim ssql As String
    Dim oconnect As Object
On Error GoTo Err
    ssql = "select t2.fitemclassid,* from t_account t1 inner join t_itemdetailv t2 on t1.fdetailid=t2.fdetailid"
    ssql = ssql & " where FAccountID=" & lAccount '& " and t2.fitemclassid=" & lType
    Set oconnect = CreateObject("K3Connection.AppConnection")
    Set rs = oconnect.Execute(ssql)
    
    '最多只允许挂客户和物料两种核算项目
    rs.Filter = "fitemclassid <> 1 and fitemclassid <> 4"
    If rs.RecordCount > 0 Then GoTo Err
    
    CheckAccountProject = True
    Set oconnect = Nothing
    Exit Function
Err:
    Set oconnect = Nothing
    CheckAccountProject = False
    Exit Function
End Function


Private Function GetCurrentArStartDate() As Date
    Dim oConn As Object
    Set oConn = CreateObject("K3Connection.AppConnection")
    Dim sql As New StringBuilder
    Dim rs As ADODB.Recordset
    
    sql.Append "set nocount on" & vbCrLf
    sql.Append "declare @year int" & vbCrLf
    sql.Append "Declare @period int" & vbCrLf
    sql.Append "Select top 1 @year = FYear , @period = FPeriod From t_subsys where Fcheckout = 0 and Fnumber = 'Ar' Order by FYear,FPeriod" & vbCrLf
    sql.Append "select FStartDate  from T_PeriodDate t1 Where FYear = @year and FPeriod = @period" & vbCrLf
    Set rs = oConn.Execute(sql.StringValue)
    If rs.EOF Then
        Err.Raise 1, "GetCurrentArStartDate", "Can't get ar min date that can input ar invoice!"
    Else
        GetCurrentArStartDate = rs.Fields("FStartDate").Value
    End If
    Set rs = Nothing
    Exit Function

End Function

Private Function GetCurrentGLStartDate() As Date
    Dim oConn As Object
    Set oConn = CreateObject("K3Connection.AppConnection")
    Dim sql As New StringBuilder
    Dim rs As ADODB.Recordset
    
    sql.Append "set nocount on" & vbCrLf
    sql.Append "declare @year int" & vbCrLf
    sql.Append "Declare @period int" & vbCrLf
    sql.Append "Select top 1 @year = FYear , @period = FPeriod From t_subsys where Fcheckout = 0 and Fnumber = 'Gl' Order by Fyear,FPeriod" & vbCrLf
    sql.Append "select FStartDate  from T_PeriodDate t1 Where FYear = @year and FPeriod = @period" & vbCrLf
    Set rs = oConn.Execute(sql.StringValue)
    If rs.EOF Then
        Err.Raise 1, "GetCurrentGLStartDate", "Can't get ar min date that can input voucher!"
    Else
        GetCurrentGLStartDate = rs.Fields("FStartDate").Value
    End If
    Set rs = Nothing
    Exit Function

End Function

Private Function ExecSql(sql As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim cnn As Object
    Set cnn = CreateObject("K3Connection.AppConnection")
    Set rs = cnn.Execute(sql)
    Set ExecSql = rs
    Set cnn = Nothing
    Set rs = Nothing
End Function

'用于防注入式攻击
Private Function SafetyStr(sqlString As String) As String
    SafetyStr = Replace(sqlString, "'", "''")
End Function

Private Function GetFileNameWithoutPath(fullfilename As String)
    Dim filenameWithoutPath As String
    Dim f() As String
    f = Split(fullfilename, "\")
    GetFileNameWithoutPath = f(UBound(f))
End Function
