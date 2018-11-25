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
'接口插件接口
Option Explicit

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private oDict As Object
Implements IBizConsole
Private Const M_VIEW = 2097152
Private m_LocalCnn As adodb.Connection
Private m_kdMain As Object


'说明:IBizConsole提供四个功能
'IBizConsole_CallFunc;IBizConsole_GetPeriodString;IBizConsole_OneTimeAlert;IBizConsole_SubID
Private Property Get LocalCnn() As adodb.Connection
    If m_LocalCnn Is Nothing Then
        MParse.ParseString MMTS.PropsString
        
        Dim dlg As Object
        Set dlg = CreateObject("CSystemDlg.Sys")
        Set m_LocalCnn = New adodb.Connection
        With m_LocalCnn
            .CursorLocation = adUseClient
            .Open dlg.LocalCnn
        End With
        Set dlg = Nothing
    End If
    Set LocalCnn = m_LocalCnn
End Property

'功能：当用户在主控台上双击某个明细功能时，主控台会创建相应的组件，并调用该组件的CallFunc方法。
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

    Select Case dctPara("FFuncName")
        Case "导出商品信息"
            Set frmObj = New frmExport
            frmObj.Show vbModal
        Case "导入包装关联"
            Set frmObj = New frmImport
            frmObj.Show vbModal
        Case "导入出库信息"
            Set frmObj = New frmImportSales
            frmObj.Show vbModal
        Case "防窜信息查询"
            Set frmObj = New frmSelect
            frmObj.Show vbModal
    End Select

    
    Set oAccess = Nothing
    
    Set frmObj = Nothing
    
    Exit Sub
Err_Exit:
    Set frmObj = Nothing
    Err.Raise Err.Number, "ExportReport.TotalCostTypeCement,IBizConsole_CallFunc" & vbCrLf & Err.Source, vbCrLf & Err.Description
End Sub


'功能：表示当前帐套期间的文本信息
Private Function IBizConsole_GetPeriodString() As String

End Function

'功能：一次性预警功能的实现代码
Private Sub IBizConsole_OneTimeAlert()
    Debug.Print "OneTimeAlert"
End Sub

'功能：主控台控件所属子系统的标识（比如gl, ar, ap, hr等）
Private Function IBizConsole_SubID() As String
    IBizConsole_SubID = SUBID
End Function

' 主控台组件所属系统的系统标识(沿用以前的系统标识)
Private Property Get SUBID() As String
    SUBID = "super"
End Property

'创建窗体
Public Function ShowFace()
    LoadBill 210000000, 1100
'    frmFilter.Show
End Function

'创建窗体
Public Function ShowExport()
'    LoadBill 210000000, 1100
    frmExport.Show vbModal
End Function

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


