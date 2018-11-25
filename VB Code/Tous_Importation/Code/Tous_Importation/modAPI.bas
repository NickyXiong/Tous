Attribute VB_Name = "API"
Option Explicit

'------------------------------------------------------------------------------------------------------------
'-------------------------------------------ȫ�ֱ�������-----------------------------------------------------
'------------------------------------------------------------------------------------------------------------

Public ValueR As Integer               ' ����򷵻�ֵ
Public Times As Integer                ' �����������
Public CancelFlag                      ' �Ƿ��˳���־
Public Pword As String                 ' �������


'------------------------------------------------------------------------------------------------------------
'-----------------------------------ʹ������ ��Ctrl+Alt+Del�� �����������---------------------------------
'------------------------------------------------------------------------------------------------------------

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" _
      (ByVal dwProcessID As Long, _
       ByVal dwType As Long _
      ) As Long
      '�˺����� API�������û��

'------------------------------------------------------------------------------------------------------------
'----------------------------------------------����λ�����ϲ�------------------------------------------------
'------------------------------------------------------------------------------------------------------------

Public Declare Function SetWindowPos Lib "user32" _
       (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long _
       ) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const Flag = SWP_NOMOVE Or SWP_NOSIZE ' ���ƶ��͸ı䴰���С


'------------------------------------------------------------------------------------------------------------
'-------------------------------------------������С������ϵͳ������-----------------------------------------
'------------------------------------------------------------------------------------------------------------

'--------------------------------------------------��������---------------------------------------------------

Public Const NIM_ADD = &H0               ' ���ָ����ͼ��
Public Const NIM_DELETE = &H2            ' ɾ��ָ����ͼ��
Public Const NIM_MODIFY = &H1            ' �޸�ָ����ͼ��
Public Const NIF_ICON = &H2              ' hIcon ������Ч
Public Const NIF_MESSAGE = &H1           ' uCallback Message ������Ч
Public Const NIF_TIP = &H4               ' szTip ������Ч
Public Const WM_MOUSEMOVE = &H200        'ʮ���� = 512   ' ����ƶ�
Public Const WM_LBUTTONDOWN = &H201      'ʮ���� = 513   ' ����������
Public Const WM_RBUTTONDOWN = &H204      'ʮ���� = 516   ' ����Ҽ�����


'--------------------------------------------------�ṹ����---------------------------------------------------

'����ϵͳ��Ҫ�ġ�������״̬������Ϣ��
Public Type NOTIFYICONDATA
        cbSize As Long                   ' �ṹ���ֽڳ���
        hwnd As Long                     ' ʹ��ͼ����ӵ�ϵͳ�����ϵ���Ϣ���մ���ľ��
        uID As Long                      ' Ӧ�ó���ָ����������ʾ��ͼ��ı�ʶ
        uFlags As Long                   ' �����޸�ͼ�ꡢ��Ϣ�͹�����ʾ�ı� �� 1.NIF_ICON    :
'                                                                              2.NIF_MESSAGE :
'                                                                              3.NIF_TIP     :
                                                                               
        uCallbackMessage As Long         ' Ӧ�ó���ָ������Ϣ��ʶ���ص���Ϣ��ֵ������ϵͳ�����IDͨ����Ϣ�����ڣ�hWnd��,��һ������¼����������е�����ͼ������Ϸ���ʱ����Щͨ�汻����
        hIcon As Long                    ' ��Ҫ��ӡ��޸Ļ�ɾ��ͼ����
        szTip As String * 64             ' ������ʾ�ı��ַ���
End Type


'---------------------------------------------------API����---------------------------------------------------

' ��ϵͳ��һ����Ϣ�Դ�ϵͳ��������ӡ��޸Ļ�ɾ��ͼ��
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" _
       (ByVal dwMessage As Long, _
        lpData As NOTIFYICONDATA _
       ) As Long

' ����ֵ ���ɹ� ���� ��0
'          ʧ�� ���� 0
' ����   ��dwMessage  : 1.NIM_ADD     : ��ϵͳ���������ͼ��
'                       2.NIM_DELETE  : ��ϵͳ������ɾ��ͼ��
'                       3.NIM_MODIFY  : ��ϵͳ�������޸�ͼ��
'          lpData     : NOTIFYICONDATA ��������

'------------------------------------------------------------------------------------------------------------
'-------------------------------------------- ��� ----------------------------------------------------------
'------------------------------------------------------------------------------------------------------------

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
       (ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long _
       ) As Long
       

'------------------------------------------------------------------------------------------------------------
'-------------------------------------------- ������Զ����� ------------------------------------------------
'------------------------------------------------------------------------------------------------------------

       ' �����ؼ���
Public Declare Function RegCloseKey Lib "advapi32.dll" _
      (ByVal hKey As Long _
      ) As Long
       ' ���ı��ַ�����ָ���Ĺؼ��ֹ���
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
      (ByVal hKey As Long, _
       ByVal lpSubKey As String, _
       phkResult As Long _
      ) As Long
      
       ' �رյ�½�ؼ���
       ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
      (ByVal hKey As Long, _
       ByVal lpValueName As String, _
       ByVal Reserved As Long, _
       ByVal dwType As Long, _
       lpData As Any, _
       ByVal cbData As Long _
      ) As Long

Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const HKEY_LOCAL_MACHINE = &H80000002

' ������Զ�����
Sub SetMyValue(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim KeyHandle&
    Dim lResult As Long
    lResult = RegCreateKey(hKey, strPath, KeyHandle&)
    lResult = RegSetValueEx(KeyHandle&, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    lResult = RegCloseKey(KeyHandle&)
End Sub


