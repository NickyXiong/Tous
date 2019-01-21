VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOpenFile1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Message"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFile 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmOpenFile1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
'CancelError 为 True。
On Error GoTo ErrHandler
    Me.Hide
    '设置过滤器。
    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls"
    '指定缺省过滤器。
    CommonDialog1.FilterIndex = 2
    '显示“打开”对话框。
    CommonDialog1.ShowOpen
    '调用打开文件的过程。
    strMappingFileName = CommonDialog1.FileName
    
    Unload Me
    Exit Sub
ErrHandler:
    '用户按“取消”按钮。
    Exit Sub
End Sub

