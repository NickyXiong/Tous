Attribute VB_Name = "MParse"
Option Explicit
Private Const cConnectString = "ConnectString"
Private Const cUserName = "UserName"
Private Const cUserID = "UserID"
Private Const cDBMSName = "DBMS Name"
Private Const cDBMSVersion = "DBMS Version"
Private m_colParse As Collection
Private m_sParseString As String
Public g_objResLoader As Object
Public g_ObjResFrmLoader As Object
Public g_Language As String
Public g_strAbnormal As String
Public g_strBeginItem As String
Public g_strBeginSEOrder As String
Public g_strBeginIcmo As String
Public g_strEndItem As String
Public g_strEndSEOrder As String
Public g_strEndIcmo As String
Public g_PreviewRs As ADODB.Recordset

Public g_strBeginMac As String
Public g_strBeginTime As String
Public g_strEndTime As String
Public g_strEndMac As String
Public g_strBeginBTime As String
Public g_strEndETime As String
Public g_lRow As Long
Public g_lItemID As Long
Public g_strMacID As String


Public Property Get PropsString() As String
    PropsString = m_sParseString
End Property
Public Property Get UserName() As String
    UserName = GetProperty(cUserName)
End Property
Public Property Get UserID() As Integer
    UserID = CInt(GetProperty(cUserID))
End Property
Public Property Get ConDBMSName() As String
    ConDBMSName = GetProperty(cDBMSName)
End Property
Public Property Get ConDBMSVersion() As String
    ConDBMSVersion = GetProperty(cDBMSVersion)
End Property
Public Property Get ConStr() As String
    ConStr = GetProperty(cConnectString)
End Property
Public Function GetProperty(ByVal sPropName As String) As String
    GetProperty = m_colParse(sPropName)
End Function
Public Function ParseString(ByVal sToParse As String) As Boolean
Dim sName As String
Dim sValue As String
    m_sParseString = sToParse
    Set m_colParse = New Collection
    Do
        sName = GetName(sToParse)
        sValue = GetValue(sToParse)
        If sName <> "" Then
            m_colParse.Add sValue, sName
        Else
            Exit Do
        End If
    Loop
    ParseString = True
End Function
Private Function SearchString(sBeSearch As String, ByVal sFind As String) As String
On Error GoTo Err_SearchString
Dim v As Variant
v = Split(sBeSearch, sFind, 2, vbTextCompare)
Dim lb As Integer, ub As Integer
    lb = LBound(v)
    ub = UBound(v)
    If ub > lb Then
        sBeSearch = v(ub)
        SearchString = v(lb)
    ElseIf ub = lb Then
        sBeSearch = ""
        SearchString = v(ub)
    Else
        sBeSearch = ""
        SearchString = ""
    End If
    Exit Function
Err_SearchString:
    sBeSearch = ""
    SearchString = ""
End Function
Private Function GetName(sBeSearch As String) As String
    GetName = SearchString(sBeSearch, "=")
    GetName = Trim$(GetName)
End Function
Private Function GetValue(sBeSearch As String) As String
sBeSearch = Trim$(sBeSearch)
If VBA.Left$(sBeSearch, 1) = "{" Then
    sBeSearch = Mid$(sBeSearch, 2)
    GetValue = SearchString(sBeSearch, "}")
    SearchString sBeSearch, ";"
Else
    GetValue = SearchString(sBeSearch, ";")
End If
    GetValue = Trim$(GetValue)
End Function

Public Function LoadKDString(ByVal strWord As String) As String '客户端弹出对话框的多语言处理
    Dim tempErrNumber As Long
    Dim tempErrDescripton As String
    Dim tempErrHelpContext As Variant
    Dim tempErrhelpfile As Variant
    Dim tempErrSource As String
    
    '临时保存传入的Err信息
    If Err.Number <> 0 Then
        tempErrDescripton = Err.Description
        tempErrHelpContext = Err.HelpContext
        tempErrhelpfile = Err.HelpFile
        tempErrSource = Err.Source
        tempErrNumber = Err.Number
    End If
    On Error GoTo HANDLEERROR
        If Len(g_Language) = 0 Then
            g_Language = GetPropertyExt("LANGUAGE")
        End If
        
        If UCase(g_Language) <> "CHS" Then
        If (g_objResLoader Is Nothing) Then
            Set g_objResLoader = CreateObject("K3RESLOADER.Loader")
        End If
        '资源文件目录
        '       g_objResLoader.ResDirectory = ""
         '资源文件本名
        g_objResLoader.ResFileBaseName = "K3INDUSTRY"
        '语言编号
        g_objResLoader.LanguageID = g_Language
        strWord = Trim(strWord)
        If InStr(1, strWord, "^|^") > 0 Then
             LoadKDString = LoadspecialKDString(strWord) '如果有特殊字符，调用LoadspecialMKDString
        Else
             LoadKDString = g_objResLoader.LoadString(Trim(strWord))
        End If
    Else
        strWord = Trim(strWord)
        If InStr(1, strWord, "^|^") > 0 Then
             LoadKDString = LoadspecialKDString(strWord) '如果有特殊字符，调用LoadspecialMKDString
        Else
            LoadKDString = strWord
        End If
    End If
    If tempErrNumber <> 0 Then
        Err.Description = tempErrDescripton
        Err.HelpContext = tempErrHelpContext
        Err.HelpFile = tempErrhelpfile
        Err.Source = tempErrSource
        Err.Number = tempErrNumber
    End If
    Exit Function
HANDLEERROR:
        LoadKDString = strWord
        If tempErrNumber <> 0 Then
            Err.Description = tempErrDescripton
            Err.HelpContext = tempErrHelpContext
            Err.HelpFile = tempErrhelpfile
            Err.Source = tempErrSource
            Err.Number = tempErrNumber
        End If
End Function

 '多语言界面初始化函数
 '参数说明：frm-传入的form或者usercontrol
 '         st-多语言化默认策略添加（默认不对ListView的ColumnHeader和comboBox内容进行多语言化）
Public Sub StrMultiLanguage(ByRef frm As Object, Optional st As KFO.Dictionary)
        Dim ResLoader As Object
        Dim Msgs As KFO.Vector
        Dim I As Long
        Dim errMessage As String
        Dim LanguageID As String
        If Len(g_Language) = 0 Then
            LanguageID = GetPropertyExt("LANGUAGE")
        Else
            LanguageID = g_Language
        End If
        If UCase(LanguageID) <> UCase("chs") Then
            Set ResLoader = CreateObject("FrmRes.FrmResLoader")
            '说明：报表窗体的listview统一处理，增加ColumnHeader默认策略
            If st Is Nothing Then
                Set st = New KFO.Dictionary
                st.Value("NeedMLColumnHeaders") = True
                Set Msgs = ResLoader.LoadFrmResStrings(frm, LanguageID, "", "K3INDUSTRY", st)
            Else '显式传入策略
                Set Msgs = ResLoader.LoadFrmResStrings(frm, LanguageID, "", "K3INDUSTRY", st)
            End If
            Set ResLoader = Nothing
        End If
        '生成异常字符串并显示，该过程按需使用
'        For i = Msgs.LBound To Msgs.UBound
'            errMessage = errMessage + Msgs(i) + vbCrLf
'        Next i
'        If errMessage <> "" Then
'            MsgBox errMessage
'        End If
End Sub

'多语言处理－模板字段替换
Public Function GetLanguageFieldName(ByVal strFieldName As String) As String
Dim strLanguageID As String

On Error GoTo H_Error
    strLanguageID = GetPropertyExt("LANGUAGE")
    Select Case UCase(strLanguageID)
        Case "CHT"
            GetLanguageFieldName = Trim(strFieldName & "_CHT")
        Case "EN"
            GetLanguageFieldName = Trim(strFieldName & "_EN")
        Case Else
            GetLanguageFieldName = Trim(strFieldName)
    End Select
    Exit Function
H_Error:
    GetLanguageFieldName = Trim(strFieldName)
End Function
Public Function GetPropertyExt(ByVal sName As String) As String
    On Error Resume Next
    Dim I As Integer
    Dim j As Integer
    Dim sTemp As String
    Dim sString As String
    Dim s As String
    sString = MMTS.PropsString
    s = ";"
    sTemp = IIf(Right(sString, 1) = s, sString, sString & s)
    sName = sName & "="
    I = InStr(1, sTemp, sName, vbTextCompare)     '不区分大小写
    If I <> 0 Then
        sTemp = Right(sTemp, Len(sTemp) - I + 1)
        j = InStr(1, sTemp, s)
        If j <> 0 Then
            sTemp = VBA.Left(sTemp, j - 1)
            GetPropertyExt = Right(sTemp, Len(sTemp) - Len(sName))
        End If
    End If
End Function
'***********************************************************
Public Function LoadspecialKDString(ByVal strWord As String) As String
     Dim Vargb As Variant
     Dim I As Long
     Dim StrCh As String
     Vargb = Split(strWord, "^|^")
     StrCh = ""
     For I = 0 To UBound(Vargb)
         If Mid(Vargb(I), 1, 3) = "~$~" Then
            Vargb(I) = Mid(Vargb(I), 4)
            StrCh = StrCh & LoadKDString(Vargb(I))
         Else
            StrCh = StrCh & Vargb(I)
         End If
     Next
     LoadspecialKDString = StrCh
End Function

Public Sub SetFormFont(frm As Object)
    If Len(g_Language) = 0 Then
        g_Language = GetPropertyExt("Language")
    End If
    If g_ObjResFrmLoader Is Nothing Then
        Set g_ObjResFrmLoader = CreateObject("FrmRes.FrmResLoader")
    End If
    g_ObjResFrmLoader.SetFrmFont frm, g_Language
End Sub

Public Function LoadKDString2(ByVal strWord As String) As String '客户端弹出对话框的多语言处理
    Dim tempErrNumber As Long
    Dim tempErrDescripton As String
    Dim tempErrHelpContext As Variant
    Dim tempErrhelpfile As Variant
    Dim tempErrSource As String
    
    '临时保存传入的Err信息
    If Err.Number <> 0 Then
        tempErrDescripton = Err.Description
        tempErrHelpContext = Err.HelpContext
        tempErrhelpfile = Err.HelpFile
        tempErrSource = Err.Source
        tempErrNumber = Err.Number
    End If
    On Error GoTo HANDLEERROR
    Dim LanguageID As String
    
    If Len(g_Language) = 0 Then
        LanguageID = GetPropertyExt("LANGUAGE")
    Else
        LanguageID = g_Language
    End If

    If UCase(LanguageID) <> UCase("chs") Then
        If (g_objResLoader Is Nothing) Then
            Set g_objResLoader = CreateObject("K3RESLOADER.Loader")
        End If
        '资源文件目录
        '       g_objResLoader.ResDirectory = ""
         '资源文件本名
        g_objResLoader.ResFileBaseName = "K3INDUSTRY"
        '语言编号
        g_objResLoader.LanguageID = LanguageID
        strWord = Trim(strWord)
        If InStr(1, strWord, "^|^") > 0 Then
             LoadKDString2 = LoadspecialKDString(strWord) '如果有特殊字符，调用LoadspecialMKDString
        Else
             LoadKDString2 = g_objResLoader.LoadString2(Trim(strWord))
        End If
    Else
        strWord = Trim(strWord)
        If InStr(1, strWord, "^|^") > 0 Then
             LoadKDString2 = LoadspecialKDString(strWord) '如果有特殊字符，调用LoadspecialMKDString
        Else
            LoadKDString2 = strWord
        End If
    End If
    If tempErrNumber <> 0 Then
        Err.Description = tempErrDescripton
        Err.HelpContext = tempErrHelpContext
        Err.HelpFile = tempErrhelpfile
        Err.Source = tempErrSource
        Err.Number = tempErrNumber
    End If
    Exit Function
HANDLEERROR:
        LoadKDString2 = strWord
        If tempErrNumber <> 0 Then
            Err.Description = tempErrDescripton
            Err.HelpContext = tempErrHelpContext
            Err.HelpFile = tempErrhelpfile
            Err.Source = tempErrSource
            Err.Number = tempErrNumber
        End If
End Function
'根据操作系统自动设置默认字体名称
Public Function SetSystemFontName() As KFO.Dictionary
    Dim ResLoader As Object
    Dim FontInfo As KFO.Dictionary
    If Len(g_Language) = 0 Then g_Language = GetPropertyExt("LANGUAGE")
    Set ResLoader = CreateObject("FrmRes.FrmResLoader")
    If Not ResLoader Is Nothing Then
        Set SetSystemFontName = ResLoader.GetK3FontProps(g_Language)
        Set ResLoader = Nothing
    End If
End Function

Public Sub SetLedgerFont(Ledger As Object)
Dim I As Long
Dim dctFont As KFO.Dictionary
Set dctFont = SetSystemFontName
With Ledger
    For I = 0 To 9
        .Printer.Header(I).Font.Name = dctFont.GetValue("Font.Name", "宋体")
        .Printer.Header(I).Font.Charset = dctFont.GetValue("Font.Charset", 134)
        .Printer.Footer(I).Font.Name = dctFont.GetValue("Font.Name", "宋体")
        .Printer.Footer(I).Font.Charset = dctFont.GetValue("Font.Charset", 134)
    Next I
End With
Set dctFont = Nothing

End Sub

Public Sub SetLedger4Font(Ledger As Object)
Dim I As Long
Dim dctFont As KFO.Dictionary
Set dctFont = SetSystemFontName
With Ledger
    For I = 1 To 8
        .TextlineFontName(I) = dctFont.GetValue("Font.Name", "宋体")
        .TextlineFontCharset(I) = dctFont.GetValue("Font.Charset", 134)
    Next I
End With
Set dctFont = Nothing
End Sub
