VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_clsPass As clsPass
Attribute m_clsPass.VB_VarHelpID = -1
Public Event Active()

Private Declare Function SetWindowPos Lib "user32" _
         (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1


Public Function PutWindowOnTop(pFrm As Form)
  Dim lngWindowPosition As Long
  
  lngWindowPosition = SetWindowPos(pFrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Function

Public Function Show(parent As Form)
    Set m_clsPass = New clsPass
'    Set m_frm = New frmProgBar
    Set m_frm.m_clsPass = m_clsPass
    m_frm.Show 0, parent
'    PutWindowOnTop m_frm
End Function


Public Function SetBarMaxValue(ByVal Value As Long)
    m_clsPass.SetBarMaxValue (Value)
End Function

Public Function SetBarMinValue(ByVal Value As Long)
    m_clsPass.SetBarMinValue (Value)
End Function

Public Function SetBarValue(ByVal Value As Long)
    m_clsPass.SetBarValue (Value)
End Function

Public Function SetMsg(ByVal Msg As String)
    m_clsPass.SetMsg (Msg)
End Function

Public Function Unload()
    m_clsPass.Unload
    Set m_frm = Nothing
    Set m_clsPass = Nothing
    DoEvents
End Function

Public Function SetBarValueWithMax()
    m_clsPass.SetBarValueWithMax
End Function
Public Function ShowProgBar()
    m_clsPass.ShowProgBar
End Function
Public Function HideProgBar()
    m_clsPass.HideProgBar
End Function


Private Sub m_clsPass_Active()
    RaiseEvent Active
End Sub


