Attribute VB_Name = "MMTS"
Option Explicit

'��ϵͳ����,�����Լ�ϵͳ�����滻
Public Const SUBID = "super"
Public Const SUBNAME = "����ϵͳ"

Private m_oSpmMgr As Object
Private m_oLogin As Object
Public LoginType As String
Public LoginAcctID As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'��¼
Public Function CheckMts(ByVal CFG As Long, Optional ByVal ChangeUser As Boolean = False) As Long
    '���Mts״̬
    CheckMts = False
    If CFG Then
        If Not m_oLogin Is Nothing And Not ChangeUser Then
           CheckMts = True
           Exit Function
        End If

        Dim bChangeMts As Boolean
        bChangeMts = True
        Set m_oLogin = CreateObject("KDLogin.clsLogin")
        If InStr(1, LoginType, "Straight", vbTextCompare) > 0 And Not ChangeUser Then
           'ֱ�ӵ���
           'ʵ�ֶ��ο���ģ������ص�¼
           If m_oLogin.LoginStraight(SUBID, SUBNAME, LoginAcctID) Then
              CheckMts = True
              Call OpenConnection
           End If
       Else
           '���µ�¼
           If m_oLogin.Login(SUBID, SUBNAME, bChangeMts) Then
              CheckMts = True
              Call OpenConnection
           End If
       End If
    Else
       m_oLogin.ShutDown
       Set m_oLogin = Nothing
    End If
End Function
'������
Private Sub OpenConnection()
    Dim lProc As Long
    lProc = GetCurrentProcessId()
    Set m_oSpmMgr = CreateObject("PropsMgr.ShareProps")
    m_oSpmMgr.addproperty lProc, "UserName", m_oLogin.UserName
    m_oSpmMgr.addproperty lProc, "PropsString", m_oLogin.PropsString
    m_oSpmMgr.addproperty lProc, "LogStatus", m_oLogin.LogStatus
    m_oSpmMgr.addproperty lProc, "AcctName", m_oLogin.AcctName
    m_oSpmMgr.addproperty lProc, "KDLogin", m_oLogin
    m_oSpmMgr.addproperty lProc, "AcctType", m_oLogin.AcctType
    m_oSpmMgr.addproperty lProc, "Setuptype", m_oLogin.SetupType
    m_oSpmMgr.addproperty lProc, "AcctID", m_oLogin.AcctID
End Sub

'��ȡ����Ϣ,�ô���Ϣ�����������Ӵ���Ϣ����������һЩ��Ϣ������μ���������Է���
Private Function GetConnectionProperty(strName As String, Optional ByVal bRaiseError As Boolean = True) As Variant
    
    Dim spmMgr As Object
    Dim lProc As Long
    lProc = GetCurrentProcessId()
    Set spmMgr = CreateObject("PropsMgr.ShareProps")
    If IsObject(spmMgr.GetProperty(lProc, strName)) Then
        Set GetConnectionProperty = spmMgr.GetProperty(lProc, strName)
    Else
        GetConnectionProperty = spmMgr.GetProperty(lProc, strName)
    End If
End Function

'------------------���Է���------------------------
'�û���
Public Function UserName() As String
If m_oLogin Is Nothing Then
    UserName = GetConnectionProperty("UserName")
Else
    UserName = m_oLogin.UserName

End If
End Function


'���Ӵ�
Public Function PropsString() As String
'for debug only
'PropsString = "ConnectString={Provider=SQLOLEDB.1;User ID=sa;Password=123;Data Source=(local);Initial Catalog=winco_hk_uat};UserName=Administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000;SubID=super;AcctType=gy;Setuptype=Industry;Language=chs;IP=169.254.244.106;MachineName=MZCSTUDIO;UUID=963564E1-EF8F-4CDA-AA3C-E9D29E3428BA"
'Exit Function
'for compile only
If m_oLogin Is Nothing Then
    PropsString = GetConnectionProperty("PropsString")
Else
    PropsString = m_oLogin.PropsString
End If
End Function
'���Ӷ���
Public Property Get ServerMgr() As Object
    Set ServerMgr = GetConnectionProperty("KDLogin")
End Property

'������
Public Function AcctName() As String
If m_oLogin Is Nothing Then
    AcctName = GetConnectionProperty("AcctName")
Else
    AcctName = m_oLogin.AcctName
End If
End Function
'------------------���Է���------------------------


Public Function ParseString() As String
    Dim var1 As Variant
    var1 = Split(PropsString, "{")
    var1 = Split(var1(1), "}")
    ParseString = var1(0)
End Function
