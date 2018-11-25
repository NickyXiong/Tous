Attribute VB_Name = "API"
Option Explicit

'------------------------------------------------------------------------------------------------------------
'-------------------------------------------全局变量声明-----------------------------------------------------
'------------------------------------------------------------------------------------------------------------

Public ValueR As Integer               ' 密码框返回值
Public Times As Integer                ' 密码输入次数
Public CancelFlag                      ' 是否退出标志
Public Pword As String                 ' 密码变量


'------------------------------------------------------------------------------------------------------------
'-----------------------------------使程序不在 “Ctrl+Alt+Del” 任务管理器中---------------------------------
'------------------------------------------------------------------------------------------------------------

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" _
      (ByVal dwProcessID As Long, _
       ByVal dwType As Long _
      ) As Long
      '此函数在 API浏览器中没有

'------------------------------------------------------------------------------------------------------------
'----------------------------------------------窗体位于最上层------------------------------------------------
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

Public Const Flag = SWP_NOMOVE Or SWP_NOSIZE ' 不移动和改变窗体大小


'------------------------------------------------------------------------------------------------------------
'-------------------------------------------程序最小化处于系统托盘中-----------------------------------------
'------------------------------------------------------------------------------------------------------------

'--------------------------------------------------常量声明---------------------------------------------------

Public Const NIM_ADD = &H0               ' 添加指定的图标
Public Const NIM_DELETE = &H2            ' 删除指定的图标
Public Const NIM_MODIFY = &H1            ' 修改指定的图标
Public Const NIF_ICON = &H2              ' hIcon 参数有效
Public Const NIF_MESSAGE = &H1           ' uCallback Message 参数有效
Public Const NIF_TIP = &H4               ' szTip 参数有效
Public Const WM_MOUSEMOVE = &H200        '十进制 = 512   ' 鼠标移动
Public Const WM_LBUTTONDOWN = &H201      '十进制 = 513   ' 鼠标左键按下
Public Const WM_RBUTTONDOWN = &H204      '十进制 = 516   ' 鼠标右键按下


'--------------------------------------------------结构声明---------------------------------------------------

'包含系统需要的“任务条状态区域信息”
Public Type NOTIFYICONDATA
        cbSize As Long                   ' 结构的字节长度
        hwnd As Long                     ' 使其图标添加到系统托盘上的消息接收窗体的句柄
        uID As Long                      ' 应用程序指定托盘上显示的图标的标识
        uFlags As Long                   ' 用于修改图标、消息和工具提示文本 ： 1.NIF_ICON    :
'                                                                              2.NIF_MESSAGE :
'                                                                              3.NIF_TIP     :
                                                                               
        uCallbackMessage As Long         ' 应用程序指定的消息标识“回调消息的值”——系统用这个ID通告消息给窗口（hWnd）,当一个鼠标事件在正在运行的托盘图标矩形上发生时，那些通告被发送
        hIcon As Long                    ' 将要添加、修改或删除图标句柄
        szTip As String * 64             ' 工具提示文本字符串
End Type


'---------------------------------------------------API声明---------------------------------------------------

' 给系统发一个消息以从系统托盘中添加、修改或删除图标
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" _
       (ByVal dwMessage As Long, _
        lpData As NOTIFYICONDATA _
       ) As Long

' 返回值 ：成功 返回 非0
'          失败 返回 0
' 参数   ：dwMessage  : 1.NIM_ADD     : 向系统托盘中添加图标
'                       2.NIM_DELETE  : 从系统托盘中删除图标
'                       3.NIM_MODIFY  : 在系统托盘中修改图标
'          lpData     : NOTIFYICONDATA 类型数据

'------------------------------------------------------------------------------------------------------------
'-------------------------------------------- 浏览 ----------------------------------------------------------
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
'-------------------------------------------- 程序的自动启动 ------------------------------------------------
'------------------------------------------------------------------------------------------------------------

       ' 建立关键字
Public Declare Function RegCloseKey Lib "advapi32.dll" _
      (ByVal hKey As Long _
      ) As Long
       ' 将文本字符串与指定的关键字关联
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
      (ByVal hKey As Long, _
       ByVal lpSubKey As String, _
       phkResult As Long _
      ) As Long
      
       ' 关闭登陆关键字
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

' 程序的自动启动
Sub SetMyValue(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim KeyHandle&
    Dim lResult As Long
    lResult = RegCreateKey(hKey, strPath, KeyHandle&)
    lResult = RegSetValueEx(KeyHandle&, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    lResult = RegCloseKey(KeyHandle&)
End Sub


