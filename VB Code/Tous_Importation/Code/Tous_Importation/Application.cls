VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�ӿڲ���ӿ�
Option Explicit

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private oDict As Object
Implements IBizConsole
Private Const M_VIEW = 2097152
Private m_LocalCnn As ADODB.Connection
Private m_kdMain As Object


'˵��:IBizConsole�ṩ�ĸ�����
'IBizConsole_CallFunc;IBizConsole_GetPeriodString;IBizConsole_OneTimeAlert;IBizConsole_SubID
Private Property Get LocalCnn() As ADODB.Connection
    If m_LocalCnn Is Nothing Then
        MParse.ParseString MMTS.PropsString
        
        Dim dlg As Object
        Set dlg = CreateObject("CSystemDlg.Sys")
        Set m_LocalCnn = New ADODB.Connection
        With m_LocalCnn
            .CursorLocation = adUseClient
            .Open dlg.LocalCnn
        End With
        Set dlg = Nothing
    End If
    Set LocalCnn = m_LocalCnn
End Property

'���ܣ����û�������̨��˫��ĳ����ϸ����ʱ������̨�ᴴ����Ӧ������������ø������CallFunc������
Private Sub IBizConsole_CallFunc(ByVal dctPara As Object, ByVal oKDMain As Object)

On Error GoTo Err_Exit
    Dim bolOK As Boolean
    Dim oAccess As Object
    Dim strUserName As String
    Dim objSyn As Object
    
    Dim frmObj As Object

'    MMTS.CheckMts True
'    modPub.Strsdsn = MMTS.PropsString
'    sdsnA = MMTS.PropsString
    strUserName = MMTS.UserName
    Set m_kdMain = oKDMain

    Select Case dctPara("FUNCID")
        Case 201607
            Set frmObj = New frmImport
            frmObj.Show vbModal
        Case "���������Ϣ"
            Set frmObj = New frmImportSales
            frmObj.Show vbModal

    End Select

    
    Set oAccess = Nothing
    
    Set frmObj = Nothing
    
    Exit Sub
Err_Exit:
    Set frmObj = Nothing
    Err.Raise Err.Number, "IBizConsole_CallFunc" & vbCrLf & Err.Source, vbCrLf & Err.Description
End Sub


'���ܣ���ʾ��ǰ�����ڼ���ı���Ϣ
Private Function IBizConsole_GetPeriodString() As String

End Function

'���ܣ�һ����Ԥ�����ܵ�ʵ�ִ���
Private Sub IBizConsole_OneTimeAlert()
    Debug.Print "OneTimeAlert"
End Sub

'���ܣ�����̨�ؼ�������ϵͳ�ı�ʶ������gl, ar, ap, hr�ȣ�
Private Function IBizConsole_SubID() As String
    IBizConsole_SubID = SUBID
End Function

' ����̨�������ϵͳ��ϵͳ��ʶ(������ǰ��ϵͳ��ʶ)
Private Property Get SUBID() As String
    SUBID = "super"
End Property

'��������
Public Function ShowFace()
    LoadBill 210000000, 1100
'    frmFilter.Show
End Function

'��������
'Public Function ShowExport()
''    LoadBill 210000000, 1100
'    frmExport.Show vbModal
'End Function

Public Function LoadBill(ByVal ClassTypeID As Long, ByVal billID As Long)
    Dim obj As Object
    Dim dataSrv As Object
    Dim dctData As Object
        
    Set dataSrv = CreateObject("K3ClassTpl.DataSrv")
    dataSrv.ClassTypeID = ClassTypeID
    
    Set obj = CreateObject("K3SingleClassBill.ClassInfoManager")
    With obj
      
        .dataSrv = dataSrv
        If IsMissing(billID) Then
            Err.Raise -1, "", "Bill ID is missing! Please contact admin user for solution."
        End If
        Set dctData = dataSrv.LoadBill(billID)
        .ISNetControl = False
        .ISCheckRight = False
        .Show billID
    End With
    Set obj = Nothing
    Set dataSrv = Nothing
    Set dctData = Nothing
End Function

