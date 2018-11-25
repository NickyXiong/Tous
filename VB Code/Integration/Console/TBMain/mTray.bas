Attribute VB_Name = "mTray"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public TrayIcon As NOTIFYICONDATA

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIF_MESSAGE = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Sub LoadIcon(ByVal bRun As Boolean)
    
    UnLoadIcon
    
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hWnd = frmMain.hWnd
    TrayIcon.uId = vbNull
    TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayIcon.ucallbackMessage = WM_MOUSEMOVE

    If bRun Then
        TrayIcon.hIcon = frmMain.imgIcon.ListImages(2).Picture
    Else
        TrayIcon.hIcon = frmMain.imgIcon.ListImages(2).Picture
    End If

    TrayIcon.szTip = "" & Chr$(0)
    
    Call Shell_NotifyIcon(NIM_ADD, TrayIcon)

End Sub

Public Sub UnLoadIcon()
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hWnd = frmMain.hWnd
    TrayIcon.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
End Sub
