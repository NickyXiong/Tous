VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportSales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "导入出库信息"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "frmImportSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5175
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
         Height          =   3570
         ItemData        =   "frmImportSales.frx":0E42
         Left            =   120
         List            =   "frmImportSales.frx":0E44
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
         Width           =   855
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
Attribute VB_Name = "frmImportSales"
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
Dim dic As KFO.Dictionary
Dim strMsg1 As String
Dim rsBill As adodb.Recordset
Dim objCreate As Object
Dim Message() As String
Dim row As Long '记录strmsg中行数
Dim BillsType As String '记录单据类型

Dim strTemp As String

On Error GoTo Err
filename = txtFile.Text
    lstResult.Clear
    '读取文件
    Set vecData = ReadExcelFile(filename)
    If vecData.UBound = 0 Then
        lstResult.AddItem "没有数据可以导入"
        lblStatus.Caption = ""
'        MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
        Exit Sub
    End If

    '将打包完的数据插入表中
    If InsertDataToTable(vecData, strMsg.StringValue) = False Then
        lstResult.AddItem "导入数据库失败"
        lblStatus.Caption = ""
'        MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
        Exit Sub
    End If

    '自动生成外购入库单
    strTemp = "select Isnull(t2.FItemID,0) FWH,Isnull(t3.FItemID,0) FDefaultWH"
    strTemp = strTemp & vbCrLf & "from T_t_Sales t1 left join t_Stock t2 on t1.FWareHouse=t2.FItemID"
    strTemp = strTemp & vbCrLf & "left join t_Stock t3 on t1.fdefaultwarehouse=t3.FItemID"
    strTemp = strTemp & vbCrLf & "Where t1.FUUID='" & strUUID & "' and t1.FType='采购订单'"
    strTemp = strTemp & vbCrLf & "and t2.FNumber not like '%C004%' and t3.FNumber not like '%C004%'"

    Set rsBill = ExecSQL(strTemp)
    If rsBill.RecordCount > 0 Then
        lblStatus.Caption = "正在生成外购入库单..."
        BillsType = "采购订单"
        Set objCreate = CreateObject("ST_M_CreateBill.ClsPurchase")
        Set dic = objCreate.CreatePurchase(MMTS.PropsString, strUUID, strMsg, BillsType)
        If dic("success") = False Then
            ExecSQL "Delete From T_t_Sales Where FUUID='" & strUUID & "' and FType='采购订单'"
            strMsg.Append dic("errmsg")
'            MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If
            
        Else
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If
        
        End If
        lblStatus.Caption = ""
    End If
    Set rsBill = Nothing
    
    
     '自动生成红字外购入库单
'    Set rsBill = ExecSQL("Select 1 from T_t_Sales Where FUUID='" & strUUID & "' and FType='采购退货'")
    strTemp = "select Isnull(t2.FItemID,0) FWH,Isnull(t3.FItemID,0) FDefaultWH"
    strTemp = strTemp & vbCrLf & "from T_t_Sales t1 left join t_Stock t2 on t1.FWareHouse=t2.FItemID"
    strTemp = strTemp & vbCrLf & "left join t_Stock t3 on t1.fdefaultwarehouse=t3.FItemID"
    strTemp = strTemp & vbCrLf & "Where t1.FUUID='" & strUUID & "' and t1.FType='采购退货'"
    strTemp = strTemp & vbCrLf & "and (t2.FNumber not like '%C004%' and t3.FNumber not like '%C004%')"

    Set rsBill = ExecSQL(strTemp)
    If rsBill.RecordCount > 0 Then
        lblStatus.Caption = "正在生成采购退货单..."
        BillsType = "采购退货"
        Set objCreate = CreateObject("ST_M_CreateBill.ClsPurchase")
        Set dic = objCreate.CreatePurchase(MMTS.PropsString, strUUID, strMsg, BillsType)
        If dic("success") = False Then
            ExecSQL "Delete From T_t_Sales Where FUUID='" & strUUID & "' and FType='采购退货'"
            strMsg.Append dic("errmsg")
'            MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If
            
        Else
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If
        
        End If
        lblStatus.Caption = ""
    End If
    Set rsBill = Nothing
    
    
    '自动生成销售出库单
'    Set rsBill = ExecSQL("Select 1 from T_t_Sales Where FUUID='" & strUUID & "' and FType='销售订单'")
    strTemp = "select Isnull(t2.FItemID,0) FWH,Isnull(t3.FItemID,0) FDefaultWH"
    strTemp = strTemp & vbCrLf & "from T_t_Sales t1 left join t_Stock t2 on t1.FWareHouse=t2.FItemID"
    strTemp = strTemp & vbCrLf & "left join t_Stock t3 on t1.fdefaultwarehouse=t3.FItemID"
    strTemp = strTemp & vbCrLf & "Where t1.FUUID='" & strUUID & "' and t1.FType='销售订单'"
    strTemp = strTemp & vbCrLf & "and (t2.FNumber not like '%C004%' and t3.FNumber not like '%C004%')"

    Set rsBill = ExecSQL(strTemp)
    If rsBill.RecordCount > 0 Then
        lblStatus.Caption = "正在生成销售出库单..."
        BillsType = "销售订单"
        Set objCreate = CreateObject("ST_M_CreateBill.ClsSalesDelievery")
        Set dic = objCreate.CreateSales(MMTS.PropsString, strUUID, strMsg, BillsType)
        If dic("success") = False Then
            ExecSQL "Delete From T_t_Sales Where FUUID='" & strUUID & "' and FType='销售订单'"
            strMsg.Append dic("errmsg")
'            MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If

        Else
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If

        End If
        lblStatus.Caption = ""
    End If
    Set rsBill = Nothing
    
        '自动生成红字销售出库单
'    Set rsBill = ExecSQL("Select 1 from T_t_Sales Where FUUID='" & strUUID & "' and FType='销售退货'")
    strTemp = "select Isnull(t2.FItemID,0) FWH,Isnull(t3.FItemID,0) FDefaultWH"
    strTemp = strTemp & vbCrLf & "from T_t_Sales t1 left join t_Stock t2 on t1.FWareHouse=t2.FItemID"
    strTemp = strTemp & vbCrLf & "left join t_Stock t3 on t1.fdefaultwarehouse=t3.FItemID"
    strTemp = strTemp & vbCrLf & "Where t1.FUUID='" & strUUID & "' and t1.FType='销售退货'"
    strTemp = strTemp & vbCrLf & "and (t2.FNumber not like '%C004%' and t3.FNumber not like '%C004%')"
    Set rsBill = ExecSQL(strTemp)
    If rsBill.RecordCount > 0 Then
        lblStatus.Caption = "正在生成销售退货单..."
        BillsType = "销售退货"
        Set objCreate = CreateObject("ST_M_CreateBill.ClsSalesDelievery")
        Set dic = objCreate.CreateSales(MMTS.PropsString, strUUID, strMsg, BillsType)
        If dic("success") = False Then
            ExecSQL "Delete From T_t_Sales Where FUUID='" & strUUID & "' and FType='销售退货'"
            strMsg.Append dic("errmsg")
'            MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If

        Else
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If

        End If
        lblStatus.Caption = ""
    End If
    Set rsBill = Nothing
    
    '自动生成其他入库单
'    Set rsBill = ExecSQL("Select 1 from T_t_Sales Where FUUID='" & strUUID & "' and FType='其它入库'")
    strTemp = "select Isnull(t2.FItemID,0) FWH,Isnull(t3.FItemID,0) FDefaultWH"
    strTemp = strTemp & vbCrLf & "from T_t_Sales t1 left join t_Stock t2 on t1.FWareHouse=t2.FItemID"
    strTemp = strTemp & vbCrLf & "left join t_Stock t3 on t1.fdefaultwarehouse=t3.FItemID"
    strTemp = strTemp & vbCrLf & "Where t1.FUUID='" & strUUID & "' and t1.FType='其它入库'"
    strTemp = strTemp & vbCrLf & "and (t2.FNumber not like '%C004%' and t3.FNumber not like '%C004%')"
    Set rsBill = ExecSQL(strTemp)
    If rsBill.RecordCount > 0 Then
        lblStatus.Caption = "正在生成其它入库单..."
        Set objCreate = CreateObject("ST_M_CreateBill.clsOtherPurchase")
        Set dic = objCreate.CreatePurchase(MMTS.PropsString, strUUID, strMsg)
        If dic("success") = False Then
            ExecSQL "Delete From T_t_Sales Where FUUID='" & strUUID & "' and FType='其它入库'"
            strMsg.Append dic("errmsg")
'            MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If
         Else
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If
           
        End If
        lblStatus.Caption = ""
        End If
        Set rsBill = Nothing
    
    
    '自动生成其他出库单
'    Set rsBill = ExecSQL("Select 1 from T_t_Sales Where FUUID='" & strUUID & "' and FType='其它出库'")
    strTemp = "select Isnull(t2.FItemID,0) FWH,Isnull(t3.FItemID,0) FDefaultWH"
    strTemp = strTemp & vbCrLf & "from T_t_Sales t1 left join t_Stock t2 on t1.FWareHouse=t2.FItemID"
    strTemp = strTemp & vbCrLf & "left join t_Stock t3 on t1.fdefaultwarehouse=t3.FItemID"
    strTemp = strTemp & vbCrLf & "Where t1.FUUID='" & strUUID & "' and t1.FType='其它出库'"
    strTemp = strTemp & vbCrLf & "and (t2.FNumber not like '%C004%' and t3.FNumber not like '%C004%')"
    Set rsBill = ExecSQL(strTemp)
    If rsBill.RecordCount > 0 Then
        lblStatus.Caption = "正在生成其它出库单..."
        Set objCreate = CreateObject("ST_M_CreateBill.clsOtherSales")
        Set dic = objCreate.CreateSales(MMTS.PropsString, strUUID, strMsg)
        If dic("success") = False Then
            ExecSQL "Delete From T_t_Sales Where FUUID='" & strUUID & "' and FType='其它出库'"
            strMsg.Append dic("errmsg")
'            MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If
        Else
            If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                Message = Split(strMsg.StringValue, vbCrLf)
                For row = LBound(Message) To UBound(Message)
                    lstResult.AddItem Message(row)
                Next
                Erase Message
                Set strMsg = Nothing
            End If
            
        End If
        lblStatus.Caption = ""
    End If
    Set rsBill = Nothing
    
    
        '自动生成调拨单
'    Set rsBill = ExecSQL("Select 1 from T_t_Sales Where FUUID='" & strUUID & "' and FType='调拨单'")
    strTemp = "select Isnull(t2.FItemID,0) FWH,Isnull(t3.FItemID,0) FDefaultWH"
    strTemp = strTemp & vbCrLf & "from T_t_Sales t1 left join t_Stock t2 on t1.FWareHouse=t2.FItemID"
    strTemp = strTemp & vbCrLf & "left join t_Stock t3 on t1.fdefaultwarehouse=t3.FItemID"
    strTemp = strTemp & vbCrLf & "Where t1.FUUID='" & strUUID & "' and t1.FType='调拨单'"
    strTemp = strTemp & vbCrLf & "and (t2.FNumber not like '%C004%' and t3.FNumber not like '%C004%')"
    Set rsBill = ExecSQL(strTemp)
    If Not rsBill Is Nothing Then
        If rsBill.RecordCount > 0 Then
            lblStatus.Caption = "正在生成调拨单..."
            Set objCreate = CreateObject("ST_M_CreateBill.clsSalesMovement")
            Set dic = objCreate.CreateSalesMovement(MMTS.PropsString, strUUID, strMsg)
            If dic("success") = False Then
                ExecSQL "Delete From T_t_Sales Where FUUID='" & strUUID & "' and FType='调拨单'"
    '            strMsg.Append dic("errmsg")
    '            MoveFile txtFile.Text, App.path & "\Failure\" & GetFileNameWithoutPath(txtFile.Text)
                If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                    Message = Split(strMsg.StringValue, vbCrLf)
                    For row = LBound(Message) To UBound(Message)
                        lstResult.AddItem Message(row)
                    Next
                    Erase Message
                    Set strMsg = Nothing
                End If
            Else
                If strMsg.StringValue <> "" Then   '将Strmsg中的信息循环打印到listbox中
                    Message = Split(strMsg.StringValue, vbCrLf)
                    For row = LBound(Message) To UBound(Message)
                        lstResult.AddItem Message(row)
                    Next
                    Erase Message
                    Set strMsg = Nothing
                End If
                
            End If
            lblStatus.Caption = ""
        End If
    End If
    Set rsBill = Nothing
    


'    MoveFile txtFile.Text, App.path & "\Imported\" & GetFileNameWithoutPath(txtFile.Text)
    lstResult.AddItem "导入结束"

    Exit Sub
Err:
    ExecSQL "Delete From T_t_Sales Where FUUID='" & strUUID & "'"
    lstResult.AddItem "导入发生错误,文件格式错误"
    lblStatus.Caption = ""

End Sub

Private Sub cmdSelected_Click()

    cmdlg.Filter = "CSV File|*.csv"
    cmdlg.FilterIndex = 1
    cmdlg.ShowOpen
    txtFile.Text = cmdlg.filename

End Sub

Private Function ReadExcelFile(filename As String) As KFO.Vector
    Dim iRow As Long
    Dim iColumn As Long
    Dim iline As Long
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlsheet As Object
    
    Dim FStockItemID As Integer
    Dim FSPID As Integer
    
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
    blnTemp = True
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
                dic("FType") = xlsheet.Cells(iRow, iCol)
            Case 2
                dic("FBillNo") = xlsheet.Cells(iRow, iCol)
                If dic("FType") = "采购订单" Then
                    If CheckPOOrder(dic("FBillNo")) = False Then
                        lstResult.AddItem "第" & CStr(iRow) & "行：采购订单 '" & dic("FBillNo") & "' 不存在"
                        blnTemp = False
                    End If
                
                ElseIf dic("FType") = "销售订单" Then
                    If CheckSEOrder(dic("FBillNo")) = False Then
                        lstResult.AddItem "第" & CStr(iRow) & "行：销售订单 '" & dic("FBillNo") & "' 不存在"
                        blnTemp = False
                    End If
                       
                End If
                
                
            Case 3
                dic("FWareHouse") = xlsheet.Cells(iRow, iCol)
                
                'Added by Nicky  - 20150422
                '如果采购订单数据的其他仓库为空，则将默认仓库的值赋给其他仓库
                If Len(dic("FWareHouse")) = 0 And dic("FType") = "采购订单" Then
                    dic("FWareHouse") = xlsheet.Cells(iRow, 7)
                End If
                
                
                If dic("FWareHouse") = "" And (dic("FType") = "采购订单" Or dic("FType") = "调拨单") Then
                    lstResult.AddItem "第" & CStr(iRow) & "行：采购订单或调拨单其他仓库不允许为空"
                    blnTemp = False
                Else
                    If dic("FType") = "采购订单" Or dic("FType") = "调拨单" Then   '只有采购订单取其他仓库，其他单据取默认仓库
        '                        If InStr(1, dic("FWareHouse"), "C004") > 0 Then   '如果仓库为C004，则不做处理
        '                            lstResult.AddItem "第" & CStr(iRow) & "行：单号 '" & dic("FBillNo") & "' 未导入，该仓库不允许手工导入"
        '                            blnTemp = False
        '                        End If
                        
                        If CheckStock(dic("FWareHouse"), FStockItemID, FSPID) = False Then
                                lstResult.AddItem "第" & CStr(iRow) & "行：仓库 '" & dic("FWareHouse") & "' 不存在"
                                blnTemp = False
                        End If
                        dic("FWareHouse") = FStockItemID
                        dic("FSPID") = FSPID
        '                    Else
        '                        If InStr(1, dic("FWareHouse"), "C004") > 0 Then   '如果仓库为C004，则不做处理
        '                            lstResult.AddItem "第" & CStr(iRow) & "行：单号 '" & dic("FBillNo") & "' 未导入，该仓库不允许手工导入"
        '                            blnTemp = False
        '                        End If
        '                        dic("FWareHouse") = 0
        '                        dic("FSPID") = 0
                    End If
                End If
            Case 4
                
                dic("FBarCode") = xlsheet.Cells(iRow, iCol)
                If dic("FBarCode") = "" Then
                    lstResult.AddItem "第" & CStr(iRow) & "行：单号 '" & dic("FBillNo") & "' 未导入，箱/盒条码不允许为空"
                    blnTemp = False
                End If
                
            Case 5
                dic("FSgin") = xlsheet.Cells(iRow, iCol)
                If dic("FSgin") = "" Then
                    lstResult.AddItem "第" & CStr(iRow) & "行：单号 '" & dic("FBillNo") & "' 未导入，条码标志不允许为空"
                    blnTemp = False
                Else
                    If CheckBarCode(dic("FSgin"), dic("FBarCode")) = False Then
                        If dic("FSgin") = 1 Then
                            lstResult.AddItem "第" & CStr(iRow) & "行：箱条码 '" & dic("FBarCode") & "' 不存在"
                        ElseIf dic("FSgin") = 0 Then
                            lstResult.AddItem "第" & CStr(iRow) & "行：盒条码 '" & dic("FBarCode") & "' 不存在"
                        End If
                        blnTemp = False
                    End If
                End If
            Case 6
               dic("FDate") = xlsheet.Cells(iRow, iCol)
               If dic("FDate") = "" Then
                    lstResult.AddItem "第" & CStr(iRow) & "行：单号 '" & dic("FBillNo") & "' 未导入，出库时间不允许为空"
                    blnTemp = False
                Else
                    '检查同一单号同一时间是否存在相同记录，存在则不允许导入
'                    If dic("FType") <> "采购退货" And dic("FType") <> "销售退货" Then
'                        If CheckBillNo(xlsheet.Cells(iRow, 2), dic("FDate")) = True Then
'                            lstResult.AddItem "第" & CStr(iRow) & "行：单据 '" & dic("FBillNo") & "' 已存在"
'                            blnTemp = False
'                        End If
'                    End If
                End If
            Case 7
                dic("FDefaultWareHouse") = xlsheet.Cells(iRow, iCol)
                If dic("FDefaultWareHouse") = "" And dic("FType") <> "采购订单" Then
                    lstResult.AddItem "第" & CStr(iRow) & "行：单号 '" & dic("FBillNo") & "' 未导入，默认仓库不允许为空"
                    blnTemp = False
                Else
                    If InStr(1, dic("FDefaultWareHouse"), "C004") > 0 Then   '如果仓库为C004，则不做处理
                        lstResult.AddItem "第" & CStr(iRow) & "行：单号 '" & dic("FBillNo") & "' 未导入，该仓库不允许手工导入"
                        blnTemp = False
                    End If
                    
                    If CheckStock(dic("FDefaultWareHouse"), FStockItemID, FSPID) = False Then
                            lstResult.AddItem "第" & CStr(iRow) & "行：仓库 '" & dic("FWareHouse") & "' 不存在"
                            blnTemp = False
                    End If
                    dic("FDefaultWareHouse") = FStockItemID
                    dic("FDefaultSPID") = FSPID
                End If
            End Select
        Next iCol
    
'    'Added by Nicky - 20150422
'    '加一个空行，区分不同行的信息，方便查看
'    lstResult.AddItem " "
    
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
        
    Set ReadExcelFile = vec
    Exit Function
HErr:
    blnTemp = False
'    Set ReadExcelFile = vec
    xlBook.Close False
    xlApp.Quit
    Set xlsheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
'    MsgBox "Error raise when importing,please check the format of " & Mid(filename, 5, Len(filename)) & Err.Description, vbInformation + vbOKOnly, "金蝶提示"
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
    Dim time As String
    
    
    '使用GUID作为一次事务的标识
    Set objTypeLib = CreateObject("Scriptlet.TypeLib")
    strUUID = CStr(objTypeLib.Guid)
    strUUID = Mid(strUUID, 1, InStr(1, strUUID, "}"))
    
    Set objTypeLib = Nothing
    
    lblStatus.Caption = "导入文件中..."
    Set objTypeLib = Nothing
On Error GoTo HErr
    InsertDataToTable = False
    ssql = ""
    ProgressBar1.Max = vctAllData.UBound
    ProgressBar1.Value = 0
    Set vctTemp = New KFO.Vector
    
    time = Format(Now, "yyyymmddhhmmss")
    
    For I = vctAllData.LBound To vctAllData.UBound
         Set dctCheck = vctAllData(I)
    
         Set dctTemp = New KFO.Dictionary
         
         
         ssql = "insert T_t_Sales "
         ssql = ssql & vbCrLf & "values('" & vctAllData(I)("FType") & "',"
         ssql = ssql & "'" & vctAllData(I)("FBillNo") & "',"
         ssql = ssql & "'" & vctAllData(I)("FWareHouse") & "',"
         ssql = ssql & "'" & vctAllData(I)("FSPID") & "',"
         ssql = ssql & "'" & vctAllData(I)("FBarCode") & "',"
         ssql = ssql & "'" & vctAllData(I)("FSgin") & "',"
         ssql = ssql & "'" & vctAllData(I)("FDate") & " ',"
         ssql = ssql & "'" & time & "',"
         ssql = ssql & "'',"
         ssql = ssql & "'',"
         ssql = ssql & "'" & strUUID & "',"
         ssql = ssql & "'',"
         ssql = ssql & "'" & vctAllData(I)("FDefaultWareHouse") & "',"
         ssql = ssql & "'" & vctAllData(I)("FDefaultSPID") & "')"
         
         dctTemp("sql") = ssql
         vctTemp.Add dctTemp
         Set dctTemp = Nothing
         ProgressBar1.Value = ProgressBar1.Value + 1
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
    lblStatus.Caption = ""
    
    Exit Function
HErr:
    InsertDataToTable = False
    If strMsg <> "" Then
       strMsg = "Following Row has be imported into DB:" & vbCrLf & strMsg
    End If
    strMsg = strMsg & vbCrLf & CNulls(Err.Description, "")
End Function

Private Function CheckBarCode(ByVal Sign As String, ByVal BarCode As String) As Boolean
Dim rs As adodb.Recordset

If Sign = 1 Then
    Set rs = ExecSQL("select 1 from t_t_package where FBoxBarCode='" & BarCode & "'")
ElseIf Sign = 0 Then
     Set rs = ExecSQL("select 1 from t_t_package where FHeBarCode='" & BarCode & "'")
End If

If rs.RecordCount = 0 Then
    CheckBarCode = False
    Exit Function
End If
CheckBarCode = True
End Function


Private Function CheckPOOrder(ByVal BillNo As String) As Boolean
Dim rs As adodb.Recordset

Set rs = ExecSQL("select 1 from POOrder where FBillNo='" & BillNo & "'")
If rs.RecordCount = 0 Then
    CheckPOOrder = False
    Exit Function
End If
CheckPOOrder = True
End Function

Private Function CheckBillNo(ByVal BillNo As String, ByVal FDate As String) As Boolean
Dim rs As adodb.Recordset

Set rs = ExecSQL("select 1 from T_t_Sales where FBillNo='" & BillNo & "' and FDate='" & FDate & "'")
If rs.RecordCount > 0 Then
    CheckBillNo = True
    Exit Function
End If
CheckBillNo = False
End Function


Private Function CheckSEOrder(ByVal BillNo As String) As Boolean
Dim rs As adodb.Recordset

Set rs = ExecSQL("select 1 from SEOrder where FBillNo='" & BillNo & "'")
If rs.RecordCount = 0 Then
    CheckSEOrder = False
    Exit Function
End If
CheckSEOrder = True


End Function

Private Function CheckStock(ByVal FNumber As String, ByRef FItemID As Integer, ByRef FSPID As Integer) As Boolean
Dim rs As adodb.Recordset
Dim str() As String
Dim strSQL As String

    If InStr(1, FNumber, "C004") > 0 Then
        str = Split(FNumber, ".")
        
        strSQL = "select t1.FNumber FStockNumber,t1.FItemID ,t2.FNumber FSPNumber,t2.FSPID "
        strSQL = strSQL & vbCrLf & "from t_Stock t1 left join t_StockPlace t2 on t1.FSPGroupID=t2.FSPGroupID "
        strSQL = strSQL & vbCrLf & "where t1.FNumber='" & str(0) & "' and t2.FNumber='" & str(1) & "'"
        
        Set rs = ExecSQL(strSQL)
        If rs.RecordCount = 0 Then
            CheckStock = False
            Exit Function
        End If
        FItemID = rs.Fields("FItemID").Value
        FSPID = rs.Fields("FSPID").Value
        CheckStock = True
    Else
        Set rs = ExecSQL("select FItemID from t_Stock where FNumber='" & FNumber & "'")
        If rs.RecordCount = 0 Then
            CheckStock = False
            Exit Function
        End If
        FItemID = rs.Fields("FItemID").Value
        FSPID = 0
        CheckStock = True
    End If

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
