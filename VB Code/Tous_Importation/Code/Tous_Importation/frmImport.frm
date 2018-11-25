VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import STN by Excel"
   ClientHeight    =   5652
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8556
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5652
   ScaleWidth      =   8556
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8292
      Begin VB.CommandButton CmdSelected 
         Height          =   300
         Left            =   6720
         Picture         =   "frmImport.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   350
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7200
         TabIndex        =   5
         Top             =   240
         Width           =   852
      End
      Begin VB.ListBox lstResult 
         Height          =   3888
         ItemData        =   "frmImport.frx":1118
         Left            =   120
         List            =   "frmImport.frx":111A
         TabIndex        =   4
         Top             =   1440
         Width           =   8000
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   8000
         _ExtentX        =   14118
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtFile 
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   5772
      End
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   8000
      End
      Begin VB.Label Label1 
         Caption         =   "File Path:"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   280
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

Private strUUID As String

Private Sub cmdImport_Click()
    Dim vecData As KFO.Vector
    Dim filename As String
    Dim strMsg As New StringBuilder
    filename = txtFile.Text

'ImportFiles filename
On Error GoTo Err:
    
    If MsgBox("Do you want to import the file to create STN(s) now?", vbYesNo, "Kingdee Prompt") = vbYes Then

        lstResult.Clear
        Set vecData = ReadExcelFile(filename)
        If vecData.UBound = 0 Then
            lstResult.AddItem "No data need to be imported."
    '            MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
            Exit Sub
        End If
        
        If InsertDataToTable(vecData, strMsg.StringValue) = False Then
            MsgBox "Import failed, please check the error log for reference.", vbCritical, "Kingdee prompt"
    '        MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
            Exit Sub
        End If
        
        CopyFile txtFile.Text, App.path & "\STN_Imported\" & GetFileNameWithoutPath(txtFile.Text)
        lstResult.AddItem "Import successfully."
    
    End If
    
    Exit Sub

Err:
    lstResult.AddItem "Import failed, please check the error log for reference."

End Sub

Private Sub cmdSelected_Click()

On Error GoTo Err

    cmdlg.Filter = "Excel File(*.xls)|*.xls"
    cmdlg.FilterIndex = 1
    cmdlg.ShowOpen
    txtFile.Text = cmdlg.filename
    
    Exit Sub
    
Err:
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
    
    Dim lItemID As Long
    
    Timestart = Now
On Error GoTo HErr
    Set xlApp = CreateObject("Excel.Application") '创建EXCEL对象
    Set xlBook = xlApp.Workbooks().Open(filename)
    Set xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表
    xlApp.Visible = False '设置EXCEL对象可见（或不可见）
    
    Dim vec As New KFO.Vector
    Dim dic As KFO.Dictionary
    
    lblStatus.Caption = "Reading Excel File..."
    
'    MMTS.CheckMts 1
'    strMsg.StringValue = ""
    lngRowsCount = xlsheet.UsedRange.Rows.Count
    lngColsCount = xlsheet.UsedRange.Columns.Count

    ProgressBar1.Max = lngRowsCount
    
    If lngColsCount <> 4 Then
    GoTo HErr
    End If
    
    blnTemp = True
    
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
                    If xlsheet.Cells(iRow, iCol) <> "" Then
                        If CheckWarehouse(Trim(xlsheet.Cells(iRow, iCol)), lItemID) = True Then
                            dic("FStockOut") = lItemID
                        Else
                            lstResult.AddItem "Row[" & iRow & "]: Stock-out Store[" & Trim(xlsheet.Cells(iRow, iCol)) & "] is not existed."
                            blnTemp = False
                        End If
                    Else
                        lstResult.AddItem "Row[" & iRow & "]: Please fill in Stock-out Store code."
                        blnTemp = False
                    End If
                 Case 2
                    If xlsheet.Cells(iRow, iCol) <> "" Then
                        If CheckWarehouse(Trim(xlsheet.Cells(iRow, iCol)), lItemID) = True Then
                            dic("FStockIn") = lItemID
                        Else
                            lstResult.AddItem "Row[" & iRow & "]: Stock-in Store[" & Trim(xlsheet.Cells(iRow, iCol)) & "] is not existed."
                            blnTemp = False
                        End If
                    Else
                        lstResult.AddItem "Row[" & iRow & "]: Please fill in Stock-out Store code."
                        blnTemp = False
                    End If
                 Case 3
                    If xlsheet.Cells(iRow, iCol) <> "" Then
                        If CheckProduct(Trim(xlsheet.Cells(iRow, iCol)), lItemID) = True Then
                            dic("FSKU") = lItemID
                        Else
                            lstResult.AddItem "Row[" & iRow & "]: SKU[" & Trim(xlsheet.Cells(iRow, iCol)) & "] is not existed."
                            blnTemp = False
                        End If
                    Else
                        lstResult.AddItem "Row[" & iRow & "]: Please fill in SKU."
                        blnTemp = False
                    End If
                     
                 Case 4
                    If xlsheet.Cells(iRow, iCol) <> "" Then
                        If IsNumeric(Trim(xlsheet.Cells(iRow, iCol))) = True And CDbl(Trim(xlsheet.Cells(iRow, iCol))) > 0 Then
                            dic("FQty") = CDbl(Trim(xlsheet.Cells(iRow, iCol)))
                        Else
                            lstResult.AddItem "Row[" & iRow & "]: Quantity should be positive number."
                            blnTemp = False
                        End If
                    Else
                        lstResult.AddItem "Row[" & iRow & "]: Please fill in Quantity."
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
    MsgBox "Data format error", vbCritical, "Kingdee Prompt"
End Function


Private Function InsertDataToTable(ByVal vctAllData As KFO.Vector, ByRef strMsg As String) As Boolean
    Dim I As Long
    Dim ssql As String
    Dim oconnect As Object
    Dim rs As ADODB.Recordset

   
    Dim strItem As String
    
    Dim objTypeLib As Object, obj As Object
    
    Dim dctCheck As KFO.Dictionary
    Dim dctTemp As KFO.Dictionary
    Dim vctTemp As KFO.Vector
    Dim strAllSQL As New StringBuilder
    
    Dim strBillNo As String
    
On Error GoTo HErr
    InsertDataToTable = False
    ssql = ""
    
    Set objTypeLib = CreateObject("Scriptlet.TypeLib")
    strUUID = CStr(objTypeLib.Guid)
    strUUID = Mid(strUUID, 1, InStr(1, strUUID, "}"))
    Set objTypeLib = Nothing
    
    lstResult.AddItem "UUID:" & strUUID

    Set vctTemp = New KFO.Vector
    
    ProgressBar1.Max = vctAllData.UBound
    lblStatus.Caption = "Importing Excel File..."
    
    For I = vctAllData.LBound To vctAllData.UBound
         Set dctCheck = vctAllData(I)
             
         Set dctTemp = New KFO.Dictionary

         
         ssql = "insert t_Tous_STNImportData (FStockOutID,FStockInID,FItemID,FQty,FUUID)"
         ssql = ssql & vbCrLf & "values('" & vctAllData(I)("FStockOut") & "',"
         ssql = ssql & "'" & vctAllData(I)("FStockIn") & "',"
         ssql = ssql & "'" & vctAllData(I)("FSKU") & "',"
         ssql = ssql & "'" & vctAllData(I)("FQty") & "',"
         ssql = ssql & "'" & strUUID & "')"
         
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
        ExecSql (strAllSQL.StringValue)
        Set oconnect = Nothing
    End If
    
    lblStatus.Caption = "Generating Stock Transfer Notices..."
    Set obj = CreateObject("Tous_M_Importation.clsImportFunction")
    If obj.CreateSTN(MMTS.PropsString, strUUID, strBillNo, strMsg) = False Then
        GoTo HErr
    End If
    Set obj = Nothing
    
    lstResult.AddItem "New Stock Transfer Notice: " & strBillNo
    lblStatus.Caption = ""
    InsertDataToTable = True
'    lstResult.AddItem "Import Success!"
  
    Exit Function
HErr:
    lblStatus.Caption = ""
    ExecSql "delete from t_Tous_STNImportData where FUUID='" & strUUID & "'"
    lstResult.AddItem strMsg
    InsertDataToTable = False
'    If strMsg <> "" Then
'       strMsg = "Following Row has be import:" & vbCrLf & strMsg
'    End If
'    strMsg = strMsg & vbCrLf & CNulls(Err.Description, "")
End Function

Private Function CheckProduct(ByVal ProductNumber As String, ByRef lItemID As Long) As Boolean
    Dim rs As ADODB.Recordset
    
    Set rs = ExecSql("select FItemID from t_ICItem where FNumber='" & ProductNumber & "'")
    If rs.RecordCount > 0 Then
        lItemID = rs.Fields("FItemID").Value
        CheckProduct = True
    Else
        CheckProduct = False
        Exit Function
    End If
End Function

Private Function CheckWarehouse(ByVal WHNumber As String, ByRef lItemID) As Boolean
    Dim rs As ADODB.Recordset
    
    Set rs = ExecSql("select FItemID from t_stock where FNumber='" & WHNumber & "'")
    If rs.RecordCount > 0 Then
        lItemID = rs.Fields("FItemID").Value
        CheckWarehouse = True
    Else
        lItemID = 0
        CheckWarehouse = False
        Exit Function
    End If
End Function

Private Function CheckBarCode(ByVal FHeBarCode As String) As Boolean
Dim rs As ADODB.Recordset

Set rs = ExecSql("select 1 from t_t_Package where FHeBarCode='" & FHeBarCode & "'")
If rs.RecordCount = 0 Then
    CheckBarCode = True
    Exit Function
End If
CheckBarCode = False
End Function

'将文件备份
Private Sub CopyFile(SourceFile As String, DestFile As String)
On Error GoTo EHandler

    Dim f As New FileSystemObject
    If f.FileExists(SourceFile) = True Then
        If f.FileExists(DestFile) = True Then
            f.DeleteFile DestFile, True
        End If
        
        SetAttr SourceFile, vbNormal
        f.CopyFile SourceFile, DestFile
    End If
    Set f = Nothing
    Exit Sub
EHandler:
    MsgBox "Copy file failed:" & Err.Description, vbCritical, "Kingdee Prompt"
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

