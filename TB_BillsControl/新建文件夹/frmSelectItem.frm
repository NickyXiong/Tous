VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmOpenFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Search"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   Icon            =   "frmOpenFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox cmbSize 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox cmbColor 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox cmbStyleNumber 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin FPUSpreadADO.fpSpread fp 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8415
      _Version        =   458752
      _ExtentX        =   14843
      _ExtentY        =   6588
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   1
      SpreadDesigner  =   "frmOpenFile.frx":0E42
   End
   Begin VB.Label Label3 
      Caption         =   "Size Code"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Color Code"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Style Number"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmOpenFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public rsAll As ADODB.Recordset

Private Sub cmbColor_Click()
    Dim strFilter As String
On Error GoTo Err
    
    If Len(cmbStyleNumber.Text) > 0 Then
        strFilter = strFilter & " and FStyleNumber='" & cmbStyleNumber.Text & "'"
    End If
    If Len(cmbColor.Text) > 0 Then
        strFilter = strFilter & " and FColorCode='" & cmbColor.Text & "'"
    End If
    If Len(cmbSize.Text) > 0 Then
        strFilter = strFilter & "and FSizeEx='" & cmbSize.Text & "'"
    End If

    If Len(strFilter) > 0 Then
        strFilter = " or (Flag=0 " & strFilter & ")"
    End If

    rsAll.Filter = "Flag=1" & strFilter
    
    rsAll.Sort = "Flag desc"
    
'    rsAll.Filter = "FColorCode='100'"
    
    fp.MaxRows = rsAll.RecordCount
    If rsAll.RecordCount > 0 Then
        rsAll.MoveFirst
        For i = 1 To rsAll.RecordCount
            With fp
                .Row = i
                .Col = 1
                .Value = rsAll.Fields("Flag").Value
                .Col = 2
                .Text = rsAll.Fields("FStyleNumber").Value
                .Lock = True
                .Col = 3
                .Text = rsAll.Fields("FColorCode").Value
                .Lock = True
                .Col = 4
                .Text = rsAll.Fields("FSizeEx").Value
                .Lock = True
                .Col = 5
                .Text = rsAll.Fields("FNumber").Value
                .Lock = True
                .Col = 6
                .Text = rsAll.Fields("FItemID").Value
                .ColHidden = True
            End With
            
            rsAll.MoveNext
        Next
    End If
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Kingdee Prompt"
End Sub

Private Sub cmbSize_Click()
    Dim strFilter As String
On Error GoTo Err
    
    If Len(cmbStyleNumber.Text) > 0 Then
        strFilter = strFilter & " and FStyleNumber='" & cmbStyleNumber.Text & "'"
    End If
    If Len(cmbColor.Text) > 0 Then
        strFilter = strFilter & " and FColorCode='" & cmbColor.Text & "'"
    End If
    If Len(cmbSize.Text) > 0 Then
        strFilter = strFilter & "and FSizeEx='" & cmbSize.Text & "'"
    End If

    If Len(strFilter) > 0 Then
        strFilter = " or (Flag=0 " & strFilter & ")"
    End If

    rsAll.Filter = "Flag=1" & strFilter
    
    rsAll.Sort = "Flag desc"
    
'    rsAll.Filter = "FColorCode='100'"
    
    fp.MaxRows = rsAll.RecordCount
    If rsAll.RecordCount > 0 Then
        rsAll.MoveFirst
        For i = 1 To rsAll.RecordCount
            With fp
                .Row = i
                .Col = 1
                .Value = rsAll.Fields("Flag").Value
                .Col = 2
                .Text = rsAll.Fields("FStyleNumber").Value
                .Lock = True
                .Col = 3
                .Text = rsAll.Fields("FColorCode").Value
                .Lock = True
                .Col = 4
                .Text = rsAll.Fields("FSizeEx").Value
                .Lock = True
                .Col = 5
                .Text = rsAll.Fields("FNumber").Value
                .Lock = True
                .Col = 6
                .Text = rsAll.Fields("FItemID").Value
                .ColHidden = True
            End With
            
            rsAll.MoveNext
        Next
    End If
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Kingdee Prompt"
End Sub

Private Sub cmbStyleNumber_Click()
    Dim strFilter As String
On Error GoTo Err
    
    If Len(cmbStyleNumber.Text) > 0 Then
        strFilter = strFilter & " and FStyleNumber='" & cmbStyleNumber.Text & "'"
    End If
    If Len(cmbColor.Text) > 0 Then
        strFilter = strFilter & " and FColorCode='" & cmbColor.Text & "'"
    End If
    If Len(cmbSize.Text) > 0 Then
        strFilter = strFilter & "and FSizeEx='" & cmbSize.Text & "'"
    End If

    If Len(strFilter) > 0 Then
        strFilter = " or (Flag=0 " & strFilter & ")"
    End If

    rsAll.Filter = "Flag=1" & strFilter
    
    rsAll.Sort = "Flag desc"
    
'    rsAll.Filter = "FColorCode='100'"
    
    fp.MaxRows = rsAll.RecordCount
    If rsAll.RecordCount > 0 Then
        rsAll.MoveFirst
        For i = 1 To rsAll.RecordCount
            With fp
                .Row = i
                .Col = 1
                .Value = rsAll.Fields("Flag").Value
                .Col = 2
                .Text = rsAll.Fields("FStyleNumber").Value
                .Lock = True
                .Col = 3
                .Text = rsAll.Fields("FColorCode").Value
                .Lock = True
                .Col = 4
                .Text = rsAll.Fields("FSizeEx").Value
                .Lock = True
                .Col = 5
                .Text = rsAll.Fields("FNumber").Value
                .Lock = True
                .Col = 6
                .Text = rsAll.Fields("FItemID").Value
                .ColHidden = True
            End With
            
            rsAll.MoveNext
        Next
    End If
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Kingdee Prompt"
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    bSelected = True
    For i = 1 To fp.MaxRows
        
    Next
End Sub

Private Sub Command2_Click()
    fp.Col = 1
    fp.Row = -1
    fp.Value = 1
End Sub

Private Sub Command3_Click()
    fp.Col = 1
    fp.Row = -1
    fp.Value = 0
End Sub

Private Sub Form_Load()
    Dim rs As Recordset
    Dim rsfield As ADODB.Field
    Dim strSQL As String
    Dim i As Integer
    Dim j As Integer
On Error GoTo Err

    MMTS.CheckMts 1
    
    strSQL = "select distinct FStyleNumber from "
    strSQL = strSQL & vbCrLf & "(select distinct t1.FStyleNumber,t2.FID as FColorCode,t1.FSizeEx from t_icitem t1"
    strSQL = strSQL & vbCrLf & "left join t_SubMessage t2 on t1.FColorExID = t2.FInterID where t2.FTypeID=11600"
    strSQL = strSQL & vbCrLf & "and isnull(t1.FStyleNumber,'')<>'' and isnull(t2.FID,'')<>'' and isnull(t1.FSizeEx,'')<>'' ) tt"
    Set rs = ExecuteSql(strSQL)
    
    cmbStyleNumber.AddItem ""
    For i = 1 To rs.RecordCount
        cmbStyleNumber.AddItem rs.Fields("FStyleNumber").Value
        rs.MoveNext
    Next
    
    strSQL = "select distinct FSizeEx from "
    strSQL = strSQL & vbCrLf & "(select distinct t1.FStyleNumber,t2.FID as FColorCode,t1.FSizeEx from t_icitem t1"
    strSQL = strSQL & vbCrLf & "left join t_SubMessage t2 on t1.FColorExID = t2.FInterID where t2.FTypeID=11600"
    strSQL = strSQL & vbCrLf & "and isnull(t1.FStyleNumber,'')<>'' and isnull(t2.FID,'')<>'' and isnull(t1.FSizeEx,'')<>'' ) tt"
    Set rs = ExecuteSql(strSQL)
    
    cmbSize.AddItem ""
    For i = 1 To rs.RecordCount
        cmbSize.AddItem rs.Fields("FSizeEx").Value
        rs.MoveNext
    Next
    
    strSQL = "select distinct FColorCode from "
    strSQL = strSQL & vbCrLf & "(select distinct t1.FStyleNumber,t2.FID as FColorCode,t1.FSizeEx from t_icitem t1"
    strSQL = strSQL & vbCrLf & "left join t_SubMessage t2 on t1.FColorExID = t2.FInterID where t2.FTypeID=11600"
    strSQL = strSQL & vbCrLf & "and isnull(t1.FStyleNumber,'')<>'' and isnull(t2.FID,'')<>'' and isnull(t1.FSizeEx,'')<>'' ) tt"
    Set rs = ExecuteSql(strSQL)
    
    cmbColor.AddItem ""
    For i = 1 To rs.RecordCount
        cmbColor.AddItem rs.Fields("FColorCode").Value
        rs.MoveNext
    Next
    
    strSQL = "select distinct t1.FItemID,t1.FStyleNumber,t2.FID as FColorCode,t1.FSizeEx,t1.FNumber,0 as Flag from t_icitem t1"
    strSQL = strSQL & vbCrLf & "left join t_SubMessage t2 on t1.FColorExID = t2.FInterID where t2.FTypeID=11600"
    strSQL = strSQL & vbCrLf & "and isnull(t1.FStyleNumber,'')<>'' and isnull(t2.FID,'')<>'' and isnull(t1.FSizeEx,'')<>''"
    Set rs = ExecuteSql(strSQL)
'    rsAll.Open strSQL, MMTS.ParseString, adOpenDynamic, adLockOptimistic
'    rsAll.LockType = adLockOptimistic
'    rsAll.CursorType = adOpenDynamic
'    Set rsAll = rs.Clone

    'rsAll包括所有的SKU数据
    Set rsAll = New ADODB.Recordset
    '构造临时记录集
    For Each rsfield In rs.Fields
        If rsfield.Type = adNumeric Then
            rsAll.Fields.Append rsfield.Name, adDecimal, rsfield.DefinedSize, adFldIsNullable
        Else
        rsAll.Fields.Append rsfield.Name, rsfield.Type, rsfield.DefinedSize, adFldIsNullable
        End If
    Next
    rsAll.CursorType = adOpenStatic
    rsAll.LockType = adLockOptimistic
    rsAll.Open
    
    If rs.RecordCount > 0 Then
        For j = 1 To rs.RecordCount
            rsAll.AddNew
            For i = 0 To rs.Fields.Count - 1
                rsAll.Fields(i).Value = rs.Fields(i).Value
            Next i
            rs.MoveNext
            rsAll.MoveLast
        Next
        rsAll.Update
    End If
    
    fp.MaxRows = rsAll.RecordCount
    rsAll.MoveFirst
    For i = 1 To rsAll.RecordCount
        With fp
            .Row = i
            .Col = 1
            .Value = rsAll.Fields("Flag").Value
            .Col = 2
            .Text = rsAll.Fields("FStyleNumber").Value
            .Lock = True
            .Col = 3
            .Text = rsAll.Fields("FColorCode").Value
            .Lock = True
            .Col = 4
            .Text = rsAll.Fields("FSizeEx").Value
            .Lock = True
            .Col = 5
            .Text = rsAll.Fields("FNumber").Value
            .Lock = True
            .Col = 6
            .Text = rsAll.Fields("FItemID").Value
            .ColHidden = True
        End With
        
        rsAll.MoveNext
    Next
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Kingdee Prompt"
End Sub



Private Sub fp_Change(ByVal Col As Long, ByVal Row As Long)
    With fp
        .Col = Col
        .Row = Row
        If .Value = 1 Then
            .Col = 6
            rsAll.Filter = "FItemID=" & CStr(.Value)
            rsAll.Update "flag", 1
'            rsAll.Fields("Flag").Value = 1
        End If
        
'        MsgBox rsAll.Fields("Flag")
        rsAll.Filter = ""
    End With
End Sub

