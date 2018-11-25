VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "导入包装关联"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton cmdImport 
         Caption         =   "确定"
         Height          =   375
         Left            =   7800
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox lstResult 
         Height          =   3960
         ItemData        =   "frmImport.frx":0E42
         Left            =   120
         List            =   "frmImport.frx":0E44
         TabIndex        =   5
         Top             =   1440
         Width           =   8775
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdSelected 
         Caption         =   ".."
         Height          =   375
         Left            =   7320
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtFile 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label lblStatus 
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
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   8775
      End
      Begin VB.Label Label1 
         Caption         =   "文件路径"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   120
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImport_Click()
    Dim vecData As KFO.Vector
    Dim filename As String
    Dim strMsg As New StringBuilder
    filename = txtFile.Text

'ImportFiles filename
On Error GoTo Err:
    lstResult.Clear
    Set vecData = ReadExcelFile(filename)
        If vecData.UBound = 0 Then
            lstResult.AddItem "没有数据可以导入"
'            MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
            Exit Sub
        End If
    
    If InsertDataToTable(vecData, strMsg.StringValue) = False Then
        MsgBox "导入失败，请查看导入日志", vbCritical, "金蝶提示"
'        MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
        Exit Sub
    End If
    
'    MoveFile txtFile.Text, App.path & "\Imported\" & GetFileNameWithoutPath(txtFile.Text)
    lstResult.AddItem "导入成功"
Exit Sub

Err:
    lstResult.AddItem "导入失败"

End Sub

Private Sub cmdSelected_Click()

    cmdlg.Filter = "CSV File|*.csv"
    cmdlg.FilterIndex = 1
    cmdlg.ShowOpen
    txtFile.Text = cmdlg.filename
    
'    txtFile.Text = GetDirectory

End Sub

Private Function ReadExcelFile(filename As String) As KFO.Vector
    Dim iRow As Long
    Dim iColumn As Long
    Dim iline As Long
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlsheet As Object
    
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
    Set xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表
    xlApp.Visible = False '设置EXCEL对象可见（或不可见）
    
    Dim vec As New KFO.Vector
    Dim dic As KFO.Dictionary
    
    lblStatus.Caption = "读取文件中..."
    
'    MMTS.CheckMts 1
'    strMsg.StringValue = ""
    lngRowsCount = xlsheet.UsedRange.Rows.Count
    lngColsCount = xlsheet.UsedRange.Columns.Count

    ProgressBar1.Max = lngRowsCount
    
    If lngColsCount <> 7 Then
    GoTo HErr
    End If
    
    Dim iCol As Long
    For iRow = 2 To lngRowsCount
        
       ProgressBar1.Value = iRow
       If Trim(xlsheet.Cells(iRow, 1)) = "" Then
           Exit For
       End If
       
       Set dic = New KFO.Dictionary
       
       For iCol = 1 To lngColsCount
       
       Select Case iCol
            Case 1
                dic("FProductNumber") = xlsheet.Cells(iRow, iCol)
                If dic("FProductNumber") <> "" Then
                    If CheckProduct(dic("FProductNumber")) = False Then
                        lstResult.AddItem "产品条码" & dic("FProductNumber") & "不存在"
                        blnTemp = False
                    End If
                Else
                    lstResult.AddItem "第 " & iRow & " 行产品条码不允许为空"
                    blnTemp = False
                End If
            Case 2
                dic("FProductName") = xlsheet.Cells(iRow, iCol)
                If dic("FProductName") = "" Then
                    lstResult.AddItem "第 " & iRow & " 行产品名称不允许为空"
                    blnTemp = False
                End If
            Case 3
                dic("FModel") = xlsheet.Cells(iRow, iCol)
                If dic("FModel") = "" Then
                    lstResult.AddItem "第 " & iRow & " 行产品规格不允许为空"
                    blnTemp = False
                End If
                
            Case 4
                dic("FProductBatch") = xlsheet.Cells(iRow, iCol)
                If dic("FProductBatch") = "" Then
                    lstResult.AddItem "第 " & iRow & " 行产品批次不允许为空"
                    blnTemp = False
                End If
            Case 5
                dic("FDate") = xlsheet.Cells(iRow, iCol)
                If dic("FDate") = "" Then
                    lstResult.AddItem "第 " & iRow & " 行到期日期不允许为空"
                    blnTemp = False
                End If
            Case 6
                dic("FBoxBarCode") = xlsheet.Cells(iRow, iCol)
                If dic("FBoxBarCode") = "" Then
                    lstResult.AddItem "第 " & iRow & " 行箱条码不允许为空"
                    blnTemp = False
                End If
            Case 7
                dic("FHeBarCode") = xlsheet.Cells(iRow, iCol)
                If dic("FHeBarCode") = "" Then
                    lstResult.AddItem "第 " & iRow & " 行盒条码不允许为空"
                    blnTemp = False
                End If
                
                If CheckBarCode(dic("FHeBarCode")) = False Then
                    lstResult.AddItem "盒条码" & dic("FHeBarCode") & "已存在"
                    blnTemp = False
                End If
        End Select
       Next iCol
      
       
        If blnTemp = False Then
            Set dic = Nothing
            blnTemp = True
            GoTo next1
        End If
        
         vec.Add dic
next1:
    Next iRow
    
    xlBook.Close False
    xlApp.Quit
    
    Set xlsheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    lblStatus.Caption = ""
    Set ReadExcelFile = vec
    Exit Function
HErr:
    blnTemp = False
    Set ReadExcelFile = vec
'    xlBook.Close False
    xlApp.Quit
    Set xlsheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    MsgBox "格式错误，请检查导入文件", vbCritical, "金蝶提示"
End Function


Private Function InsertDataToTable(ByVal vctAllData As KFO.Vector, ByRef strMsg As String) As Boolean
    Dim I As Long
    Dim ssql As String
    Dim oconnect As Object
    Dim rs As adodb.Recordset

   
    Dim strItem As String
    
    Dim objTypeLib As Object
    
    Dim dctCheck As KFO.Dictionary
    Dim dctTemp As KFO.Dictionary
    Dim vctTemp As KFO.Vector
    Dim strAllSQL As New StringBuilder
    
    
    
    Set objTypeLib = Nothing
On Error GoTo HErr
    InsertDataToTable = False
    ssql = ""
    
    Set vctTemp = New KFO.Vector
    
    ProgressBar1.Max = vctAllData.UBound
    lblStatus.Caption = "Importing File..."
    
    For I = vctAllData.LBound To vctAllData.UBound
         Set dctCheck = vctAllData(I)
             
         Set dctTemp = New KFO.Dictionary
           
         
         ssql = "insert T_t_Package "
         ssql = ssql & vbCrLf & "values('" & vctAllData(I)("FProductNumber") & "',"
         ssql = ssql & "'" & vctAllData(I)("FProductName") & "',"
         ssql = ssql & "'" & vctAllData(I)("FModel") & "',"
         ssql = ssql & "'" & vctAllData(I)("FProductBatch") & "',"
         ssql = ssql & "'" & vctAllData(I)("FDate") & "',"
         ssql = ssql & "'" & vctAllData(I)("FBoxBarCode") & "',"
         ssql = ssql & "'" & vctAllData(I)("FHeBarCode") & "')"
         
         dctTemp("sql") = ssql
         vctTemp.Add dctTemp
         Set dctTemp = Nothing
         ProgressBar1.Value = I
    Next
    
    If strMsg <> "" Then
       GoTo HErr
    Else
        'CmdImport.Enabled = True
    End If
    
    strAllSQL.Append "set nocount on"
    
    For I = vctTemp.LBound To vctTemp.UBound
        strAllSQL.Append vbCrLf & vctTemp(I)("sql")
        If I Mod 50 = 0 Then
       '    Debug.Print strAllSQL
           Set oconnect = CreateObject("K3Connection.AppConnection")
            oconnect.Execute (strAllSQL.StringValue)
            Set oconnect = Nothing
            strAllSQL.Remove 1, Len(strAllSQL.StringValue)
            strAllSQL.Append "set nocount on"
        End If
    Next


    If strAllSQL.StringValue <> "set nocount on" Then
      ' Debug.Print strAllSQL        Set oconnect = CreateObject("K3Connection.AppConnection")
        ExecSQL (strAllSQL.StringValue)
        Set oconnect = Nothing
    End If
    InsertDataToTable = True
    lstResult.AddItem "Import Success!"
  
    Exit Function
HErr:
    InsertDataToTable = False
    If strMsg <> "" Then
       strMsg = "Following Row has be import:" & vbCrLf & strMsg
    End If
    strMsg = strMsg & vbCrLf & CNulls(Err.Description, "")
End Function

Private Function CheckProduct(ByVal ProductNumber As String) As Boolean
Dim rs As adodb.Recordset

Set rs = ExecSQL("select 1 from t_ICItem where FBarCode='" & ProductNumber & "'")
If rs.RecordCount = 0 Then
    CheckProduct = False
    Exit Function
End If
CheckProduct = True
End Function

Private Function CheckBarCode(ByVal FHeBarCode As String) As Boolean
Dim rs As adodb.Recordset

Set rs = ExecSQL("select 1 from t_t_Package where FHeBarCode='" & FHeBarCode & "'")
If rs.RecordCount = 0 Then
    CheckBarCode = True
    Exit Function
End If
CheckBarCode = False
End Function

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
    MsgBox "Move file failed:" & Err.Description, vbOKOnly, "金蝶提示"
    Err.Clear
End Sub

Private Function GetFileNameWithoutPath(fullfilename As String)
    Dim filenameWithoutPath As String
    Dim f() As String
    f = Split(fullfilename, "\")
    GetFileNameWithoutPath = f(UBound(f))
End Function

Private Sub ImportFiles(ByVal strSourcePath As String)
    Dim strFileName As String
    Dim arrFileList() As String '用于存放需要导入的文件名称
    Dim strMsg As New StringBuilder
    Dim I As Integer
    Dim vecData As KFO.Vector
    
    If Len(strSourcePath) <> 0 Then
        '读取导入目录下的所有csv文件列表
        I = 0
        ReDim Preserve arrFileList(I) As String

        strFileName = Dir(strSourcePath & "\*.csv")
        Do While strFileName <> ""
            If UCase(Right(strFileName, 3)) = "CSV" Then
                I = I + 1
                ReDim Preserve arrFileList(I) As String
                arrFileList(I) = strFileName
            End If
            strFileName = Dir() '读取下一个文件
        Loop

        If UBound(arrFileList) = 0 Then
            Exit Sub
        End If

        '开始逐文件导入
        For I = 1 To UBound(arrFileList)
'            Call Sleep(50)
            strFileName = Trim(arrFileList(I))
            

            Set vecData = ReadExcelFile(strSourcePath & "\" & strFileName)
           
          If InsertDataToTable(vecData, strMsg.StringValue) = False Then
            MsgBox "Import Failed!", vbCritical, "金蝶提示"
'            MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
            Exit Sub
          End If
          
'            MoveFile txtFile.Text, App.path & "\Imported\" & GetFileNameWithoutPath(txtFile.Text)
            'MsgBox "Import Success!", vbOKOnly, "金蝶提示"
        Next
    Else
        MsgBox "Folder path must be entered", vbInformation, "金蝶提示"
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
