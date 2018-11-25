VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLisence 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lisence制作工作"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   5640
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   360
      Left            =   5640
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "打开Lisence"
      Height          =   360
      Left            =   5640
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "生成Lisence"
      Height          =   360
      Left            =   5640
      TabIndex        =   6
      Top             =   720
      Width           =   1350
   End
   Begin VB.TextBox txtInfo 
      Height          =   975
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmLisence.frx":0000
      Top             =   3600
      Width           =   3735
   End
   Begin VB.TextBox txtTask 
      Height          =   2775
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox txtCompany 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "说        明："
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "流程名称："
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公司名称："
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frmLisence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONST_KEY As String = "FUN4.ORG All Rights Reserved!"
Private Const CONST_CAPTION As String = "FUN4"

Private Sub cmdInput_Click()
    Dim sText As String
    Dim sKey As String
    Dim sPath As String
    
    Dim c As Lisence
    Dim l As Long
    Dim v() As String
    Dim b() As Byte
On Error GoTo HERROR
    
    l = FreeFile
    sPath = GetOpenString
    Open sPath For Input As #l
    Input #l, sText
    Close #l
    
    sText = Encrypt(sText, 188, 24)
    v = Split(sText, vbCrLf)
    If UBound(v) > 2 Then
        If v(0) <> CONST_CAPTION Then
            Err.Raise -1, "不是有效的Lisence文件！"
        End If
        
        txtCompany.text = v(1)
        txtTask.text = v(2)
    Else
        Err.Raise -1, "不是有效的Lisence文件！"
    End If
    
    Erase v
    MsgBox "OK!", vbInformation & vbOKOnly
    
    Exit Sub
    
HERROR:
    MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub cmdOutput_Click()
    Dim sText As String
    Dim sKey As String
    Dim sPath As String
    
    Dim l As Long
On Error GoTo HERROR

    sText = CONST_CAPTION & vbCrLf & txtCompany.text & vbCrLf & txtTask
    sText = Encrypt(sText, 188, 24)
    
    l = FreeFile
    sPath = GetSaveString
    If Len(sPath) > 0 Then
        Open sPath For Output As #l
        Print #l, sText
        Close #l
    End If
    
    MsgBox "OK!", vbInformation & vbOKOnly
    
    Exit Sub
    
HERROR:
    MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Function GetSaveString() As String
    FileDialog.FileName = ""
    FileDialog.Filter = "*.lic|*.lic"
    FileDialog.ShowSave
    
    GetSaveString = FileDialog.FileName
End Function

Private Function GetOpenString() As String
    FileDialog.FileName = ""
    FileDialog.Filter = "*.lic|*.lic"
    FileDialog.ShowSave
    
    GetOpenString = FileDialog.FileName
End Function


Private Function Encrypt(ByVal strSource As String, ByVal Key1 As Byte, _
ByVal Key2 As Integer) As String
Dim bLowData As Byte
Dim bHigData As Byte
Dim i As Integer
Dim strEncrypt As String
Dim strChar As String
For i = 1 To Len(strSource)
strChar = Mid(strSource, i, 1)
bLowData = AscB(MidB(strChar, 1, 1)) Xor Key1
bHigData = AscB(MidB(strChar, 2, 1)) Xor Key2
strEncrypt = strEncrypt & ChrB(bLowData) & ChrB(bHigData)
Next
Encrypt = strEncrypt
End Function
