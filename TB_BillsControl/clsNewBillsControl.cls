VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewBillsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is BillEvent Interface Class, made by K3BOSPLUGINSWIZAED"
 
'定义 BillEvent 接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillInterface  As BillEvent
Attribute m_BillInterface.VB_VarHelpID = -1

Private strCustProduct As String
Private strCustomer As String
Private strProduct As String

Private strContact As String
Private strPaymentTerm As String
 
Private bIsSelBill As Boolean
 
Public Sub Show(ByVal oBillInterface As Object)
 
'    MMTS.CheckMts 1
    'BillEvent 接口实现
    '注意: 此方法必须存在, 请勿修改
   Set m_BillInterface = oBillInterface
    

End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = Nothing

End Sub




Private Function GetKeyField(strFieldName As String) As String
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim strFilter As String
Dim I As Long
On Error GoTo Err

    strSQL = "select FKey from ICClassTableInfo where FClassTypeID =" & m_BillInterface.FID & " and FCaption_CHS='" & strFieldName & "'"
    Set rs = modPub.ExecuteSql(strSQL)
    
    If rs.RecordCount > 0 Then
        GetKeyField = rs.Fields("FKey").Value
    End If

    Set rs = Nothing
    Exit Function
Err:
    Set rs = Nothing
    MsgBox "GetKeyField error:" & Err.Description, vbInformation, "Kingdee Prompt"
End Function


Private Function GetMaxID(strID As String) As String
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim strFilter As String
Dim I As Long
On Error GoTo Err

    strSQL = "select max(cast(substring(fbillno,10,4) as int)) + 10001 as MaxID from t_BOS200000001 where LEFT(fbillno,5)='" & strID & Format(Now, "yyyy") & "'"
    Set rs = modPub.ExecuteSql(strSQL)
    
    If IsNull(rs.Fields("MaxID").Value) = True Then
        GetMaxID = strID & Format(Now, "yyyymmdd") & "0001"
    Else
        GetMaxID = strID & Format(Now, "yyyymmdd") & Right(CStr(rs.Fields("MaxID").Value), 4)
    End If

    Set rs = Nothing
    Exit Function
Err:
    Set rs = Nothing
    MsgBox "GetMaxID error:" & Err.Description, vbInformation, "Kingdee Prompt"
End Function


Private Sub m_BillInterface_MenuBarClick(ByVal BOSTool As K3ClassEvents.BOSTool, Cancel As Boolean)
    Dim I As Long
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim lCol As Long
    Dim Row As Long
    Dim dct As KFO.Dictionary
    Dim frm As Object
    
    Dim fs As New FileSystemObject
    Dim vec As KFO.Vector
On Error GoTo Err
    Select Case BOSTool.ToolName
        '相关历史记录查询
        Case "mnuImport"
            Set frm = New frmOpenFile1
            frm.Show vbModal
            
            If fs.FileExists(strMappingFileName) = True Then
                Set vec = ReadExcelFile(strMappingFileName)
                FillSelectedData vec
'                ImportMapping strMappingFileName
            Else
                If Len(strMappingFileName) > 0 Then
                    MsgBox "File dosen't exist!", vbCritical, "Kingdee Prompt"
                End If
            End If
         Case "mnuProductSearch"
            Set frm = New frmOpenFile
            frm.Show vbModal
            If frm.bSelected = True Then
                Set rs = frm.rsAll
                Set vec = SelectItems(rs)
                FillSelectedData vec
            End If
         Case "mnuResolve"
         
            If CStr(m_BillInterface.GetFieldValue("FCheckerID", -1, Enu_ValueType_FFLD)) <> "" Then
 
                Set frm = New frmResolved
                frm.lngInterID = CStr(m_BillInterface.GetFieldValue("FID", -1, Enu_ValueType_FFLD))
                frm.m_strRemark = CStr(m_BillInterface.GetFieldValue("FResolvedRemark", -1, Enu_ValueType_FFLD))
                frm.m_strStatus = CStr(m_BillInterface.GetFieldValue("FResolvedStatus", -1, Enu_ValueType_FFLD))
                frm.m_bNewBill = True
                frm.Show vbModal
                
                If frm.m_bUpdated = True Then
                    If CStr(m_BillInterface.GetFieldValue("FResolvedStatus", -1, Enu_ValueType_FFLD)) <> "Y" Then
                        m_BillInterface.SetFieldValue "FResolvedRemark", frm.m_strRemark, -1
                        m_BillInterface.SetFieldValue "FResolvedStatus", "Y", -1
                    Else
                        m_BillInterface.SetFieldValue "FResolvedRemark", frm.m_strRemark, -1
                        m_BillInterface.SetFieldValue "FResolvedStatus", "N", -1
                    End If
                End If

            End If
            
    End Select
    
    Exit Sub
Err:
    Set rs = Nothing
    MsgBox Err.Description, vbInformation, "Kingdee Prompt"
End Sub

Private Sub m_BillInterface_MenuBarInitialize(ByVal oMenuBar As K3ClassEvents.MenuBar)
    Select Case m_BillInterface.FID
        Case 200000101
            Dim oTool   As K3ClassEvents.BOSTool
            Dim oBand   As K3ClassEvents.BOSBand
            Set oTool = oMenuBar.BOSTools.Add("mnuImport")
            With oTool
                .Caption = "Upload"
                .ToolTipText = "Upload Excel"
                .Description = "Upload"
                .ShortcutKey = 0
                .Visible = True
                .Enabled = True
                .BeginGroup = True
                .ToolPicture = GetAppPath() & "FOLDER03.ico"
                .SetPicture 0, vbButtonFace
            End With
            Set oBand = oMenuBar.BOSBands("BandToolBar")
            oBand.BOSTools.InsertAfter "mnuCaculate", oTool
            
            Set oTool = oMenuBar.BOSTools.Add("mnuProductSearch")
            With oTool
                .Caption = "Search"
                .ToolTipText = "Product Search"
                .Description = "Product Search"
                .ShortcutKey = 0
                .Visible = True
                .Enabled = True
                .BeginGroup = True
                .ToolPicture = GetAppPath() & "TASKL.ico"
                .SetPicture 0, vbButtonFace
            End With
            Set oBand = oMenuBar.BOSBands("BandToolBar")
            oBand.BOSTools.InsertAfter "mnuCaculate", oTool
            
            Set oTool = oMenuBar.BOSTools.Add("mnuResolve")
            With oTool
                .Caption = "Resolve"
                .ToolTipText = "Resolve"
                .Description = "Resolve"
                .ShortcutKey = 0
                .Visible = True
                .Enabled = True
                .BeginGroup = True
                .ToolPicture = GetAppPath() & "KEY04.ICO"
                .SetPicture 0, vbButtonFace
            End With
            Set oBand = oMenuBar.BOSBands("BandToolBar")
            oBand.BOSTools.InsertAfter "mnuCaculate", oTool
    End Select
End Sub
Public Function GetAppPath() As String
    GetAppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
End Function

'Import data from excel file
'strFileName is as Excel file name
Function ReadExcelFile(strFileName) As KFO.Vector
    Dim I, Row As Long
    
    Dim xlApp As Excel.Application '定义EXCEL类
    Dim xlBook As Excel.Workbook '定义工件簿类
    Dim xlsheet As Excel.Worksheet '定义工作表类
    
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application") '创建EXCEL对象
    Set xlBook = xlApp.Workbooks().Open(strFileName)
    Set xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表
    xlApp.Visible = False

    Dim vec As New Vector
    Dim dic As KFO.Dictionary
    
    I = 2
    Row = 1
    m_BillInterface.DeleteEntryData 3
    Do While Len(xlsheet.Cells(I, 1)) > 0
        Set dic = New KFO.Dictionary
        dic("FStyleNumber") = Trim(xlsheet.Cells(I, 1))
        dic("FColor") = Trim(xlsheet.Cells(I, 2))
        dic("FSize") = Trim(xlsheet.Cells(I, 3))
        dic("FQty") = Trim(xlsheet.Cells(I, 4))
        vec.Add dic
        I = I + 1
    Loop
    xlBook.Close False
    xlApp.Quit
    Set xlsheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    Set ReadExcelFile = vec
End Function

Function SelectItems(rs As ADODB.Recordset) As KFO.Vector
    Dim vec As New KFO.Vector
    Dim dic As KFO.Dictionary
    
    rs.Filter = "Flag=1"
    rs.MoveFirst
    While rs.EOF = False
        Set dic = New KFO.Dictionary
        If rs.Fields("Flag") = 1 Then
            dic("FStyleNumber") = rs.Fields("FStyleNumber").Value
            dic("FColor") = rs.Fields("FColorCode").Value
            dic("FSize") = rs.Fields("FSizeEx").Value
            dic("FItemNumber") = rs.Fields("FNumber").Value
            dic("FQty") = 0
            vec.Add dic
        End If
        rs.MoveNext
    Wend
    Set SelectItems = vec
End Function

Sub FillSelectedData(vec As KFO.Vector)
    Dim errStr As New StringBuilder
    Dim dic As KFO.Dictionary
    Dim sql As String
    'Find material number
    Dim rs As ADODB.Recordset
    For I = 1 To vec.Size
        Set dic = vec(I)
        If dic.Lookup("FItemNumber") = False Then
            sql = "Select t1.FNumber From t_icitem t1 Where isnull(t1.FStyleNumber,'')='" _
                & vec(I)("FStyleNumber") & "' and isnull(t1.FColorEx,'')='" & vec(I)("FColor") & "' and isnull(t1.FSizeEx,'')='" & vec(I)("FSize") & "'"
            Set rs = modPub.ExecuteSql(sql)
            If rs.RecordCount = 1 Then
                vec(I)("FItemNumber") = rs.Fields("FNumber").Value
            ElseIf rs.RecordCount > 1 Then
                errStr.Append "The reocrd in " & (I + 1) & "row:Style Number:" & dic("FStyleNumber") & " and Color:" & dic("FColor") & " and Size:" & dic("FSize") & " record has more than one records" & Chr(10)
                vec(I)("FItemNumber") = ""
            ElseIf rs.RecordCount = 0 Then
                errStr.Append "The reocrd in " & (I + 1) & "row:Style Number:" & dic("FStyleNumber") & " and Color:" & dic("FColor") & " and Size:" & dic("FSize") & " not mapping to material" & Chr(10)
                vec(I)("FItemNumber") = ""
            End If
            
            Set rs = Nothing
        End If
    Next
    
    
    I = 2
    Row = 1
    m_BillInterface.DeleteEntryData 3
    For I = 1 To vec.Size
        If vec(I)("FItemNumber") <> "" Then
            m_BillInterface.SetFieldValue "FItemID", vec(I)("FItemNumber"), Row
            m_BillInterface.SetFieldValue "FExpQty", vec(I)("FQty"), Row
            m_BillInterface.InsertNewRowAndFill 3, Row + 1
          '  m_BillInterface.InsertNewRowAndFill 3, Row, "FItemID", vec(I)("FItemNumber") ', "FExpQty", vec(I)("FQty")
            Row = Row + 1
        End If
    Next
    
    m_BillInterface.RemoveRow , 3, Row
    
    If errStr.Length > 0 Then
        Shell "notepad.exe ", vbNormalFocus
        SendKeys errStr.StringValue
    End If
    
End Sub
'
'Sub ImportMapping(strFileName)
'Dim strSQL As String
'Dim rsAccount As ADODB.Recordset
'Dim strErr As String
'Dim i, Row As Long
'
'Dim xlApp As Excel.Application '定义EXCEL类
'Dim xlBook As Excel.Workbook '定义工件簿类
'Dim xlsheet As Excel.Worksheet '定义工作表类
'
'On Error Resume Next
'    Set xlApp = CreateObject("Excel.Application") '创建EXCEL对象
'    Set xlBook = xlApp.Workbooks().Open(strFileName)
'    Set xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表
'
'    xlApp.Visible = False
'
''    strSQL = "select FNumber,FName from t_account where FDetail=1"
''    Set rsAccount = modPub.ExecuteSql(strSQL)
'
'    Dim errStr As New StringBuilder
'    Dim vec As New Vector
'    Dim dic As KFO.Dictionary
'
'    i = 2
'    Row = 1
'    m_BillInterface.DeleteEntryData 3
'    Do While Len(xlsheet.Cells(i, 1)) > 0
'        Set dic = New KFO.Dictionary
'        dic("FStyleNumber") = Trim(xlsheet.Cells(i, 1))
'        dic("FColor") = Trim(xlsheet.Cells(i, 2))
'        dic("FSize") = Trim(xlsheet.Cells(i, 3))
'        dic("FQty") = Trim(xlsheet.Cells(i, 4))
'        vec.Add dic
'        i = i + 1
'    Loop
'    xlBook.Close False
'    xlApp.Quit
'    Set xlsheet = Nothing
'    Set xlBook = Nothing
'    Set xlApp = Nothing
'
'    Dim sql As String
'    'Find material number
'    Dim rs As ADODB.Recordset
'    For i = 1 To vec.Size
'        sql = "Select t1.FNumber From t_icitem t1 left join t_SubMessage t2 on t1.FColorExID=t2.FInterID Where isnull(t1.FStyleNumber,'')='" _
'            & vec(i)("FStyleNumber") & "' and isnull(t2.FID,'')='" & vec(i)("FColor") & "' and isnull(t1.FSizeEx,'')='" & vec(i)("FSize") & "'"
'        Set rs = modPub.ExecuteSql(sql)
'        If rs.RecordCount = 1 Then
'            vec(i)("FItemNumber") = rs.Fields("FNumber").Value
'        ElseIf rs.RecordCount > 1 Then
'            errStr.Append "The reocrd in " & (i + 1) & "row:Style Number:" & dic("FStyleNumber") & " and Color:" & dic("FColor") & " and Size:" & dic("FSize") & " record has more than one records" & Chr(10)
'            vec(i)("FItemNumber") = ""
'        ElseIf rs.RecordCount = 0 Then
'            errStr.Append "The reocrd in " & (i + 1) & "row:Style Number:" & dic("FStyleNumber") & " and Color:" & dic("FColor") & " and Size:" & dic("FSize") & " not mapping to material" & Chr(10)
'            vec(i)("FItemNumber") = ""
'        End If
'        Set rs = Nothing
'    Next
'
'
'    i = 2
'    Row = 1
'    m_BillInterface.DeleteEntryData 3
'    For i = 1 To vec.Size
'        If vec(i)("FItemNumber") <> "" Then
'            m_BillInterface.SetFieldValue "FItemID", vec(i)("FItemNumber"), Row
'            m_BillInterface.SetFieldValue "FExpQty", vec(i)("FQty"), Row
'            m_BillInterface.InsertNewRowAndFill 3, Row + 1
'          '  m_BillInterface.InsertNewRowAndFill 3, Row, "FItemID", vec(I)("FItemNumber") ', "FExpQty", vec(I)("FQty")
'            Row = Row + 1
'        End If
'    Next
'
'    If errStr.Length > 0 Then
'        Shell "notepad.exe ", vbNormalFocus
'        SendKeys errStr.StringValue
''        For I = 1 To UBound(s())
''            SendKeys s(I)
''        Next I
'    End If
''    Do While Len(xlsheet.Cells(I, 1)) > 0
''
''
''        If rsAccount.RecordCount > 0 Then
''            m_BillInterface.InsertNewRowAndFill 2, Row, "FAccountNumberInK3", rsAccount.Fields("FNumber").Value, _
''                                                            "FAccountNameInK3", rsAccount.Fields("FName").Value, _
''                                                            "FAccountNumberInOrbit", Trim(xlsheet.Cells(I, 3)), _
''                                                            "FAccountNameInOrbit", Trim(xlsheet.Cells(I, 4)), _
''                                                            "FChsDesc", Trim(xlsheet.Cells(I, 5)), _
''                                                            "FEngDesc", Trim(xlsheet.Cells(I, 6))
''            Row = Row + 1
''        Else
''            strErr = strErr & "Account " & Trim(xlsheet.Cells(I, 1)) & " doesn't exist in K/3!" & vbCrLf
''        End If
''
''        I = I + 1
''    Loop
''
''    If Len(strErr) > 0 Then
''
''    End If
''    Exit Sub
'
'End Sub
