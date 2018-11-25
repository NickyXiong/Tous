VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "导出商品信息"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   5280
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   ".."
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Top             =   210
      Width           =   300
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   210
      Width           =   4575
   End
   Begin VB.CommandButton btnExport 
      Caption         =   "导出"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblCaution 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "导出路径"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExport_Click()
    Dim ssql As String
    Dim oconnect As Object
    Dim rsItem As ADODB.Recordset
    Dim I As Long
    Dim j As Integer
    Dim strPrint As New StringBuilder
    Dim strEntry As String
    Dim filenumber As Long
    Dim filename As String
    Dim f() As String
    
On Error GoTo Err
                
    lblCaution.Caption = "正在导出，请稍后"
    filename = txtFile.Text
    
    ssql = "select FNumber,FName,FModel from t_ICItem"
    Set rsItem = ExecSQL(ssql)
    If rsItem.RecordCount > 0 Then
        strPrint.Append "产品编号,产品名称,产品规格" & vbCrLf
        For I = 1 To rsItem.RecordCount
            For j = 0 To 2
                If j <> 2 Then
                    strPrint.Append Chr(34) & "=" & Chr(34) & Chr(34) & rsItem.Fields(j).Value & Chr(34) & Chr(34) & Chr(34) & ","
                Else
                     strPrint.Append Chr(34) & "=" & Chr(34) & Chr(34) & rsItem.Fields(j).Value & Chr(34) & Chr(34) & Chr(34) & vbCrLf
                End If
            Next
            rsItem.MoveNext
        Next
        strPrint.Append strEntry
        filenumber = FreeFile
        Open filename For Output As #filenumber
        Print #filenumber, strPrint.StringValue
        Close #filenumber
        MsgBox "导出成功！", vbInformation, "金蝶提示"
    Else
        MsgBox "导出失败！", vbInformation, "金蝶提示"
    End If
    
    lblCaution.Caption = ""
    Exit Sub
Err:
    MsgBox "导出失败：" & Err.Description, vbInformation, "金蝶提示"
    lblCaution.Caption = ""

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()

    cmdlg.Filter = "csv file|*.csv"
    cmdlg.ShowOpen
    
    txtFile.Text = cmdlg.filename

End Sub
