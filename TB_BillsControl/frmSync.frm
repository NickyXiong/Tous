VERSION 5.00
Begin VB.Form frmSync 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private m_IsOk As Boolean
'Private m_strCellList As String
'Public m_rsTemplate As ADODB.Recordset
'Dim strList As String '字段名称
'
'Public Property Get IsOk() As Boolean
'    IsOk = m_IsOk
'End Property
'
'Public Function GetFilter() As KFO.Dictionary
'    Set GetFilter = ScriptScheme
'End Function
'
'
'Private Sub cmdCancel_Click()
'    m_IsOk = False
'    Me.Hide
'End Sub
'
'Private Sub cmdOK_Click()
'    On Error GoTo HERROR
'    If DTPickerFrm.Value > DTPickerTo.Value Then
'        MsgBox "起始日期不能大於截止日期!", vbInformation, "金蝶提示!"
'        Exit Sub
'    End If
'
'    If Not CheckScheme Then Exit Sub
'    If LvwScheme.SelectedItem.Text = "默認方案" Then
'        MParse.ParseString MMTS.PropsString
'        mPublic.SaveRptScheme MParse.UserID, 100015, "默認方案", ScriptScheme()
'    End If
'    m_IsOk = True
'    Me.Hide
'    Exit Sub
'HERROR:
'    MsgBox Err.Description, vbInformation, "金蝶提示"
'End Sub
'
'
'
'
'Private Sub Form_Initialize()
'    m_IsOk = False
'End Sub
'
'
'Private Sub Form_Load()
'
'    InitScheme
'
'End Sub
'
'
'
'Private Sub LvwScheme_ItemClick(ByVal item As MSComctlLib.ListItem)
'    SetThisScheme LvwScheme.SelectedItem.Text
'End Sub
'
'
'
'
'
'Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'    ButtonClick Button.Key
'End Sub
'
'Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'    ButtonMenuClick ButtonMenu.Key
'End Sub
'
'
'
''////////////////////////////////////////////基础资料查看/////////////////////////////////////////////////'
'
'Private Function SetBaseNumber(ByVal aItemClassID As Long, Optional ByVal KeyWord As String, Optional ByVal Filter As String) As String
'    '添加错误陷阱
'    On Error GoTo HERROR
'
'    Dim retObj As Object
'    Set retObj = GLView.ItemLookup(aItemClassID, KeyWord, Filter)
'    If retObj.returnok Then SetBaseNumber = retObj.ReturnObject.Number
'
'    Set retObj = Nothing
'    Exit Function
'
'HERROR:
'    MsgBox Err.Description
'End Function
'
''////////////////////////////////////////////辅助资料查看/////////////////////////////////////////////////'
'
'Private Function SetSubMessage(ByVal TypeID As Long, Optional ByVal KeyWord As String, Optional ByVal Filter As String)
'    '添加错误陷阱
'    On Error GoTo HERROR
'    Dim retObj As Object
'
'    Set retObj = GLView.SubMesLookup(TypeID, KeyWord, Filter)
'    If retObj.returnok Then SetSubMessage = retObj.ReturnObject.Number
'
'    Set retObj = Nothing
'    Exit Function
'HERROR:
'    MsgBox Err.Description
'End Function
'
'Private Sub ButtonClick(ButtonKey As String)
'    Select Case UCase(ButtonKey)
'    Case "SAVE"
'        SaveScheme
'    Case "DELETE"
'        DeleteScheme
'    Case "CLEAR"
'        ClearScheme
'    End Select
''    Debug.Print ButtonKey
'End Sub
'
'Private Sub ButtonMenuClick(ByVal ButtonMenuKey As String)
'    Select Case UCase(ButtonMenuKey)
'    Case "LARGEICON"
'        LvwScheme.View = lvwIcon
'    Case "SMALLICON"
'        LvwScheme.View = lvwSmallIcon
'    Case "LIST"
'        LvwScheme.View = lvwList
'    Case "DETAIL"
'        LvwScheme.View = lvwReport
'    End Select
'    Debug.Print ButtonMenuKey
'End Sub
'
'Private Sub SaveScheme()
'    Dim SchemeName As String
'    Dim dict As KFO.Dictionary
'    Dim I As Long
'    Dim bHas As Boolean
'
'    On Error GoTo HERROR
'    If LvwScheme.SelectedItem.Text = "默認方案" Then
'        SchemeName = Trim(InputBox("请输入方案名：", "保存方案"))
'    Else
'        SchemeName = LvwScheme.SelectedItem.Text
'    End If
'    If Trim(SchemeName) = "" Then Exit Sub
'
'    Set dict = ScriptScheme()
'
'    MParse.ParseString MMTS.PropsString
'    mPublic.SaveRptScheme MParse.UserID, 100015, SchemeName, dict
'
'    For I = 1 To LvwScheme.ListItems.Count
'        If UCase(Trim(LvwScheme.ListItems.item(I).Text)) = UCase(Trim(SchemeName)) Then
'            LvwScheme.ListItems(I).Selected = True
'            SetThisScheme SchemeName
'            bHas = True
'            Exit For
'        End If
'    Next I
'
'    If Not bHas Then
'        LvwScheme.ListItems.Add LvwScheme.ListItems.Count + 1, "Key" & SchemeName, SchemeName, 1
'    End If
'
'    Set dict = Nothing
'    Exit Sub
'HERROR:
'    Set dict = Nothing
'    MsgBox "保存方案错误。" & Err.Description, vbInformation, "金蝶提示"
'End Sub
'
'Private Function ScriptScheme() As KFO.Dictionary
'    Dim dict As KFO.Dictionary
'    Dim I As Long
'    Dim selectitem As String
'    Dim item(1 To 100) As Variant
'    Dim svs As Variant
'    Dim strAutoFilter As String
'    Dim lRow As Long
'    Dim lCol As Long
'    Dim lMaxRow As Long '最大行
'    Dim strFiledName As String
'    Set dict = New KFO.Dictionary
'
'    dict("DateFrom") = Format(DTPickerFrm.Value, "yyyy-mm-dd")
'    dict("DateTo") = Format(DTPickerTo.Value, "yyyy-mm-dd")
'
'
'
'    Set ScriptScheme = dict
'    Set dict = Nothing
'End Function
'
'Private Sub DeleteScheme()
'
'    Dim SchemeName As String
'
'    On Error GoTo HERROR
'    SchemeName = LvwScheme.SelectedItem.Text
'
'    MParse.ParseString MMTS.PropsString
'    mPublic.DelRptScheme MParse.UserID, 100015, SchemeName
'    ClearScheme
'
'    If SchemeName <> "默認方案" Then
'        LvwScheme.ListItems.Remove LvwScheme.SelectedItem.Index
'    End If
'    Exit Sub
'HERROR:
'    MsgBox "删除方案错误。" & Err.Description, vbInformation, "金蝶提示"
'End Sub
'
'Private Sub ClearScheme()
'
'
'
'End Sub
'
'Private Sub InitScheme()
'    Dim rs As ADODB.Recordset
'    Dim I As Long
'
'    On Error GoTo HERROR
'
'    MParse.ParseString MMTS.PropsString
'    Set rs = GetAllScheme(MParse.UserID, 100015)
'
'    LvwScheme.ListItems.Clear
'    LvwScheme.ListItems.Add 1, "默認方案", "默認方案", 1
'
'    I = 2
'    While Not rs.EOF
'        If rs.Fields("FPlan") <> "默認方案" Then
'            LvwScheme.ListItems.Add I, "Key" & rs.Fields("FPlan"), rs.Fields("FPlan"), 1
'            I = I + 1
'        End If
'        rs.MoveNext
'    Wend
'
'    LvwScheme.ListItems(1).Selected = True
'
'
'
'
'
'    ClearScheme
'    SetThisScheme LvwScheme.ListItems(1).Text
'
'
'    Set rs = Nothing
'    Exit Sub
'HERROR:
'    Set rs = Nothing
'    MsgBox "装载方案错误。" & Err.Description, vbInformation, "金蝶提示"
'End Sub
'
'
'Private Sub SetThisScheme(ByVal SchemeName As String)
'
'    Dim rs As ADODB.Recordset
'    Dim selectitem As String
'    Dim I As Long
'    Dim item() As String
'    Dim sStock As String
'    Dim svs As Variant
'    Dim dict As KFO.Dictionary
'    Dim strAutoFilter As String
'    Dim strArray() As String
''    Dim i As Long
'    Dim lRow As Long
'    Dim lCol As Long
'    On Error Resume Next
'
'    MParse.ParseString MMTS.PropsString
'    Set rs = GetThisScheme(MParse.UserID, 100015, SchemeName)
'    Set dict = New KFO.Dictionary
'
'
'    rs.Filter = "FKey='DateFrom'"
'    DTPickerFrm.Value = rs.Fields("fvalue")
'    rs.Filter = "FKey='DateTo'"
'    DTPickerTo.Value = rs.Fields("fvalue")
'
'
'
'    Set rs = Nothing
''    Exit Sub
''HError:
''    Set rs = Nothing
''    MsgBox "装载方案错误。" & Err.Description, vbInformation, "金蝶提示"
'
'End Sub
'
'Private Function CheckScheme() As Boolean
'    CheckScheme = True
'End Function
'
'
'
'
'Public Function GetAllScheme(ByVal UserID As Long, _
'                            ByVal RptID As Long) As ADODB.Recordset
'
'    Dim strSql As String
'
'    strSql = "Select Distinct FPlan From ICReportProfile " & _
'            "Where FRptID=" & RptID & _
'            " AND FUserID=" & UserID
'
'    Set GetAllScheme = mPublic.ExecuteSql(strSql)
'
'End Function
'Public Function GetThisScheme(ByVal UserID As Long, _
'                            ByVal RptID As Long, _
'                            ByVal PlanName As String) As ADODB.Recordset
'
'    Dim strSql As String
'
'    strSql = "Select FPlan,FKey,FValue From ICReportProfile " & _
'            "Where FRptID=" & RptID & _
'            " AND FUserID=" & UserID & _
'            " AND FPlan='" & GetQuoted(PlanName) & "'"
'
'    Set GetThisScheme = mPublic.ExecuteSql(strSql)
'
'End Function
'
'
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
''    If KeyCode <> vbKeyF7 Then Exit Sub
''    Dim iItemClassID As Long
''    Select Case Me.ActiveControl.hWnd
''        Case txtSaleChanelNameFrom.hWnd, txtSaleChanelNameTo.hWnd
''            If cmbSaleChanel.ListIndex = 0 Then
''                 Me.ActiveControl.Text = SetBaseNumber(2)
''            ElseIf cmbSaleChanel.ListIndex = 1 Then
''                Me.ActiveControl.Text = SetBaseNumber(1)
''            Else
''
''            End If
''
''    End Select
'End Sub
'
'
'
'Private Function SetCurrceny() As String
'    '添加错误陷阱
'    On Error GoTo HERROR
'    Dim oRtn As Object
'
'
'
'    Set oRtn = GLView.CurrencyLookup()
'            If Not (oRtn Is Nothing) Then
''                Set retDict = New KFO.Dictionary
''                retDict.Value("InterID") = oRtn.CurrencyID
'                If oRtn.returnok Then
'                    SetCurrceny = oRtn.ReturnObject.Number
'                End If
''                retDict.Value("Name") = oRtn.Name
'                Set oRtn = Nothing
'            End If
''
'
'    Exit Function
'HERROR:
'    MsgBox Err.Description
'End Function
'
'
'Private Function SetUserItem() As String
'    Dim obj As Object
'    Dim rs As ADODB.Recordset
'    Dim lItemID As Long
'    Dim sNumber As String
'    Dim sName As String
'On Error GoTo hrr
'    Set obj = CreateObject("K3BaseList.BaseList")
'    MParse.ParseString MMTS.PropsString
'    Set rs = obj.Show(MParse.UserID, -5)
'    If Not rs Is Nothing Then
'        If rs.RecordCount > 0 Then
'            lItemID = rs.Fields("FUserID")
'            sNumber = rs.Fields("fName")
'            sName = rs.Fields("fName")
'        End If
'    End If
'    SetUserItem = sName
'    Set rs = Nothing
'    Set obj = Nothing
'    Exit Function
'hrr:
'    MsgBox "获取用户资料出错,原因可能是:" & Err.Description, vbInformation, "金蝶提示!"
'    Set rs = Nothing
'    Set obj = Nothing
'End Function
'Private Function SetBosBaseItem(ByVal lItemclassid As Long) As String
'    Dim oDataSrv As Object
'    Dim obj As Object
'    Dim retObj As Object
'    Dim lItemID As Long
'    Dim sNumber As String
'    Dim sName As String
'On Error GoTo hrr
'    Set oDataSrv = CreateObject("K3ClassTpl.DataSrv")
'    oDataSrv.ClasstypeID = lItemclassid
'    Set obj = CreateObject("K3ClassLookup.SingleBillLookup")
'    obj.ClasstypeID = lItemclassid
'    obj.ShowType = 1
'    Set retObj = obj.Show
'    If retObj Is Nothing Then
''        ViewItem = False
'        Set oDataSrv = Nothing
'        Set obj = Nothing
'        Exit Function
'    Else
'        lItemID = retObj(1)("fid")
'        sNumber = retObj(1)("fNumber")
'        sName = retObj(1)("fName")
'    End If
'    SetBosBaseItem = sNumber
'    Set oDataSrv = Nothing
'    Set obj = Nothing
'    Exit Function
'hrr:
'    MsgBox "获取用户资料出错,原因可能是:" & Err.Description, vbInformation, "金蝶提示!"
'    Set oDataSrv = Nothing
'    Set obj = Nothing
'End Function
'
'
'
'
