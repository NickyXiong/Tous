Attribute VB_Name = "MMTS"
Option Explicit
'子系统描述,根据自己系统内容替换
Public SUBID As String
Public Const SUBNAME = "Super"
Public m_LanguageType As String
'mts share property lockmode
Private Const LockMethod = 1
Private Const LockSetGet = 0
'mts share property
Private Const Process = 1
Private Const Standard = 0
Public LoginType As String
Public LoginAcctID As Long
'Private m_oSvrMgr As Object 'Server Manager
Private m_oSpmMgr As Object
Private m_oLogin As Object
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Function CheckMts(CFG As Long) As Long
   '检查Mts状态
'''    CheckMts = False
'''    If CFG Then
'''        Dim bChangeMts As Boolean
'''        bChangeMts = CanChangeMtsServer()
'''        Set m_oLogin = Nothing
'''        Set m_oLogin = CreateObject("KDLogin.clsLogin")
'''        If m_oLogin.Login(SUBID, SUBNAME, bChangeMts) Then
'''            CheckMts = True
'''            Call OpenConnection
'''        End If
'''    Else
'''       m_oLogin.ShutDown
'''       Set m_oLogin = Nothing
'''    End If
    CheckMts = False
    If CFG Then
        Dim bFirst As Boolean
        If m_oLogin Is Nothing Then
           bFirst = True
        End If

        Dim bChangeMts As Boolean
        bChangeMts = True
        Set m_oLogin = Nothing
        Set m_oLogin = CreateObject("KDLogin.clsLogin")
        If InStr(1, LoginType, "Straight", vbTextCompare) > 0 And bFirst Then
           If m_oLogin.LoginStraight(SUBID, SUBNAME, LoginAcctID) Then
              CheckMts = True
              Call OpenConnection
           End If
       Else
'           If m_oLogin.login(SUBID, SUBNAME, bChangeMts) Then
           If m_oLogin.login("super", "***K3外挂***", True) Then
              CheckMts = True
              Call OpenConnection
           End If
       End If
    Else
       m_oLogin.ShutDown
       Set m_oLogin = Nothing
    End If

End Function
Public Function UserName() As String
If m_oLogin Is Nothing Then
    UserName = GetConnectionProperty("UserName")
Else
    UserName = m_oLogin.UserName
End If
End Function
Public Function PropsString() As String
If m_oLogin Is Nothing Then
    PropsString = GetConnectionProperty("PropsString")
Else
    PropsString = m_oLogin.PropsString
End If
End Function
Public Property Get ServerMgr() As Object
    Set ServerMgr = GetConnectionProperty("KDLogin")
End Property
Public Function IsDemo() As Boolean
If m_oLogin Is Nothing Then
    IsDemo = (GetConnectionProperty("LogStatus") = 2)
Else
    IsDemo = (m_oLogin.LogStatus = 2)
End If
End Function
Public Function AcctName() As String
If m_oLogin Is Nothing Then
    AcctName = GetConnectionProperty("AcctName")
Else
    AcctName = m_oLogin.AcctName
End If
End Function
Public Function AcctID() As String
If m_oLogin Is Nothing Then
    AcctID = GetConnectionProperty("AcctID")
Else
    AcctID = m_oLogin.AcctID
End If
End Function
Private Function GetConnectionProperty(strName As String, Optional ByVal bRaiseError As Boolean = True) As Variant
    
    Dim spmMgr As Object
    'Dim spmGroup As Object
    'Dim spmProp As Object
    'Dim bExists As Boolean
    
    'Set spmMgr = CreateObject("MTxSpm.SharedPropertyGroupManager.1")
    'Set spmGroup = spmMgr.CreatePropertyGroup("Info", LockSetGet, Process, bExists)
    
    'Set spmProp = spmGroup.Property(strName)
    'If IsObject(spmProp.Value) Then
    '    Set GetConnectionProperty = spmProp.Value
    'Else
    '    GetConnectionProperty = spmProp.Value
    'End If
    Dim lProc As Long
    lProc = GetCurrentProcessId()
    Set spmMgr = CreateObject("PropsMgr.ShareProps")
    If IsObject(spmMgr.GetProperty(lProc, strName)) Then
        Set GetConnectionProperty = spmMgr.GetProperty(lProc, strName)
    Else
        GetConnectionProperty = spmMgr.GetProperty(lProc, strName)
    End If
End Function
Private Sub OpenConnection()
    'Dim spmMgr As Object
    'Dim spmGroup As Object
    'Dim spmProp As Object
    'Dim bExists As Boolean
    
    'Set spmMgr = CreateObject("MTxSpm.SharedPropertyGroupManager.1")
    'Set spmGroup = spmMgr.CreatePropertyGroup("Info", LockSetGet, Process, bExists)
    'Set spmProp = spmGroup.CreateProperty("UserName", bExists)
    'spmProp.Value = m_oLogin.UserName
    'Set spmProp = spmGroup.CreateProperty("PropsString", bExists)
    'spmProp.Value = m_oLogin.PropsString
    'Set spmProp = spmGroup.CreateProperty("KDLogin", bExists)
    'spmProp.Value = m_oLogin
    Dim lProc As Long
    lProc = GetCurrentProcessId()
    Set m_oSpmMgr = CreateObject("PropsMgr.ShareProps")
    m_oSpmMgr.addproperty lProc, "UserName", m_oLogin.UserName
    m_oSpmMgr.addproperty lProc, "PropsString", m_oLogin.PropsString
    m_oSpmMgr.addproperty lProc, "LogStatus", m_oLogin.LogStatus
    m_oSpmMgr.addproperty lProc, "AcctName", m_oLogin.AcctName
    m_oSpmMgr.addproperty lProc, "AcctID", m_oLogin.AcctID
    m_oSpmMgr.addproperty lProc, "KDLogin", m_oLogin
End Sub
Private Sub CloseConnection()
On Error Resume Next
Dim lProc As Long
    lProc = GetCurrentProcessId()
    m_oSpmMgr.delproperty lProc, "UserName"
    m_oSpmMgr.delproperty lProc, "PropsString"
    m_oSpmMgr.delproperty lProc, "LogStatus"
    m_oSpmMgr.delproperty lProc, "AcctName"
    m_oSpmMgr.delproperty lProc, "AcctID"
    m_oSpmMgr.delproperty lProc, "KDLogin"
    Set m_oSpmMgr = Nothing
End Sub
Public Function IsIndustry() As Boolean
    IsIndustry = (UCase(GetConnectionProperty("AcctType")) = "GY")
End Function

Public Function GetPropertyExt(ByVal sName As String) As String
    
    On Error Resume Next
    Dim I As Integer
    Dim j As Integer
    Dim sTemp As String
    Dim sString As String
    Dim s As String
    
    sString = PropsString
    s = ";"
    
    sTemp = IIf(Right(sString, 1) = s, sString, sString & s)
    sName = sName & "="
    
    I = InStr(1, sTemp, sName, vbTextCompare)     '不区分大小写
    If I <> 0 Then
        sTemp = Right(sTemp, Len(sTemp) - I + 1)
        j = InStr(1, sTemp, s)
        If j <> 0 Then
            sTemp = VBA.Left(sTemp, j - 1)
            GetPropertyExt = UCase$(Right(sTemp, Len(sTemp) - Len(sName)))
        End If
    End If
End Function

Public Function LoadString(ByVal MesIndex As Long) As String
    m_LanguageType = GetPropertyExt("Language")
    If UCase(m_LanguageType) = UCase("CHS") Then
        LoadString = LoadResString(Val(MesIndex & "2"))
    ElseIf UCase(m_LanguageType) = UCase("CHT") Then
        LoadString = LoadResString(Val(MesIndex & "0"))
    ElseIf UCase(m_LanguageType) = UCase("EN") Then
        LoadString = LoadResString(Val(MesIndex & "1"))
    End If
End Function


Public Function UserID() As String
    Dim strProps As String
    Dim I As Long
    Dim vUserID
    Dim vValue
    strProps = PropsString
    I = InStr(1, strProps, "UserID=", vbTextCompare)
    If I > 0 Then
        strProps = Right(strProps, Len(strProps) - I + 1)
        vUserID = Split(strProps, ";")
        vValue = Right(vUserID(0), Len(vUserID(0)) - Len("UserID="))
        UserID = vValue
    End If
    
End Function

Public Function ParseString() As String
    Dim var1 As Variant
    var1 = Split(PropsString, "{")
    var1 = Split(var1(1), "}")
    ParseString = var1(0)
End Function

