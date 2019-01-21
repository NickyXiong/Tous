VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOldBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_LocalCnn As ADODB.Connection

Public Function OpenList(ByVal ListID As Long _
            , Optional ByVal bShowFilter As Boolean = False _
            , Optional ByVal sFilter As String = "")
10        On Error GoTo errBack
          
          Dim frm As frmMain
20        Set frm = New frmMain
30        If ListID <> 0 Then
40            With frm.ICList
50              MParse.ParseString MMTS.PropsString
60                .BaseInfo = MParse.ConStr & "|" & MParse.UserID & "|" & MParse.UserName
70                .ListID = ListID
80                .ListShowStatus = IIf(bShowFilter, 0, 3) '0 有下推 3无过滤界面 ,3多编辑 2\6.选单 8.查看模式
90                .LocalCnn = LocalCnn.ConnectionString
100               .MultiSelect = True
110               .CanSelectMultiRow = True
120               .Filter = sFilter
130               .SelectFlag = IIf(bShowFilter, 1, 0)
140               If .InitData Then
150                   .SetOtherInfo
160                   .ToolBarVis = True
170                   frm.Caption = .ListTemplateInfo.Item(1).Value("FName") & "序时薄"
180               Else
190                   Set frm = Nothing
200                   Exit Function
210               End If
220           End With
230       End If
240       frm.Show vbModal
250       Set frm = Nothing
260       Exit Function
errBack:
270       MsgBox Err.Description & " OpenList Class Module.clsOldBill" & "(" & Erl() & ")", _
                 vbOKOnly + vbInformation, "金蝶提示"
End Function

Private Function NewBill(ByVal nBillCls As Long, ByVal nBillTranType As String)
10        On Error GoTo errBack
          
          Dim objTemp As Object
          
20        Set objTemp = CreateObject("K3Bills.Bills")
30        With objTemp
35             MParse.ParseString MMTS.PropsString
40            .LocalCnn = LocalCnn.ConnectionString
50            .SystemName = "UN"
60            .SetOpt MParse.UserID, MParse.UserName
70            If nBillTranType <> "" Then
80                .NewBillTransType = nBillTranType
90            End If
100           .Show nBillCls, 0
110       End With
120       Set objTemp = Nothing
130       Exit Function
errBack:
140       MsgBox Err.Description & " OpenList Class Module.clsOldBill" & "(" & Erl() & ")", _
                 vbOKOnly + vbInformation, "金蝶提示"
End Function

'Public Function ViewBill(ByVal Lcnnstr As String, _
'                          ByVal SystemName As String, _
'                          ByVal UserID As Long, _
'                          ByVal UserName As String, _
'                          ByVal BillCls As Enu_BillCls, _
'                          ByVal BillTransType As Long, _
'                          ByVal BillInterID As Long, _
'                          ByVal ShowBillType As Enu_ShowBillType)
'    Dim objTemp As Object
'    Dim FieldDescVect As KFO.Vector
'    Dim tdict As KFO.Dictionary
'    On Error GoTo HError
'    If BillInterID = 0 Then
'        MsgBox LoadKDString("无单据信息（内码为0）。"), vbOKOnly, LoadKDString("金蝶提示")
'        Exit Function
'    End If
'
'    'KDLOGMSG_BEGIN
'    Dim LogI__0 As Long
'    LogI__0 = LogStart("K3Bills.frmBill.ViewBill 预览")
'    'KDLOGMSG_END
'    Set FieldDescVect = New KFO.Vector
'    Set tdict = New KFO.Dictionary
'    tdict.Value("FColName") = "FInterID"
'    tdict.Value("FISPrimary") = 1
'    FieldDescVect.Add tdict
'    Set tdict = New KFO.Dictionary
'    tdict.Value("FColName") = "FTranType"
'    tdict.Value("FISPrimary") = 3
'    FieldDescVect.Add tdict
'    Set tdict = Nothing
'    Set objTemp = CreateObject("K3Bills.Bills")
'    With objTemp
'        .LocalCnn = Lcnnstr
'        .SystemName = SystemName
'        .SetOpt UserID, UserName
'        .ListRecordset = GetBillRs(BillTransType, BillInterID)
'        .ListRSFieldVect = FieldDescVect
'        .Show BillCls, ShowBillType
'    End With
'    Set objTemp = Nothing
'    Set FieldDescVect = Nothing
'    'KDLOGMSG_BEGIN
'    LogEnd LogI__0, 0
'    'KDLOGMSG_END
'    Exit Function
'HError:
'    MsgBox LoadKDString("数据读取错误。") & Err.Description, vbOKOnly, LoadKDString("金蝶提示")
'    'KDLOGMSG_BEGIN
'    LogEnd LogI__0, 1
'    'KDLOGMSG_END
'End Function
'                ViewBill LCnn, gSystemName, UserID, UserName, NewBillCls, NewBillTransType, RelateBillInterID, NewBillShowType



Private Function GetBillRs(ByVal BillTransType As Long, _
                           ByVal lInterID As Long) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    With rs.Fields
        .Append "FInterID", adInteger
        .Append "FTranType", adInteger
    End With
    rs.Open
    rs.AddNew
    rs.Fields("FInterID") = lInterID
    rs.Fields("FTranType") = BillTransType
    rs.Update
    Set GetBillRs = rs
    Set rs = Nothing
End Function

Public Function LookUpList(ByVal ListID As Long _
            , Optional ByVal bShowFilter As Boolean = False _
            , Optional ByVal sFilter As String = "") As Object
          Dim frm As New frmMain
          Dim obj As Object
          Dim rtObj As KFO.Vector, rDict As KFO.Dictionary
          
10        On Error GoTo errBack
20        MParse.ParseString MMTS.PropsString

30        Set frm = New frmMain
40        If ListID <> 0 Then
50            Set frm = New frmMain
60                With frm.ICList
70                    .ListID = ListID
80                    .BaseInfo = MParse.ConStr & "|" & MParse.UserID & "|" & MParse.UserName
90                    .LocalCnn = LocalCnn.ConnectionString
100                   .ToolBarVis = True
110                   .Filter = sFilter
120                   .SelectFlag = IIf(bShowFilter, 1, 0)
130                   .CanSelectMultiRow = True
140                   .MultiSelect = 1
150                   .ListShowStatus = 2
160                   If .InitData Then
170                       .SetOtherInfo
180                       frm.Caption = .ListTemplateInfo.Item(1).Value("FName") & "序时薄"
190                   Else
200                       Set frm = Nothing
210                       Exit Function
220                   End If
230               End With
240       End If
250       frm.Show vbModal
260       If Not frm.ICList.OK Then
270           If Not frmMain Is Nothing Then
280               Unload frmMain
290               Set frmMain = Nothing
300           End If
310           Exit Function
320       End If
330       Set rtObj = frm.ICList.GetSelected()
340       If rtObj Is Nothing Then
350           If Not frmMain Is Nothing Then
360               Unload frmMain
370               Set frmMain = Nothing
380           End If
390           Exit Function
400       End If
          
410       If rtObj.UBound = 0 Then
420           If Not frm.ICList Is Nothing Then
430               Unload frm.ICList
440              Set frmMain = Nothing
450           End If
460           Exit Function
470       End If
480       Set LookUpList = rtObj
toExit:
490       Set frm = Nothing
500       Set rtObj = Nothing
510       Set rDict = Nothing
520       Set obj = Nothing
530       Exit Function
errBack:
540       MsgBox Err.Description & " LookUpList_Class Module.clsOldBill" & "(" & Erl() & ")", _
                 vbOKOnly + vbInformation, "金蝶提示"
End Function

'说明:IBizConsole提供四个功能'Geelong Edit By xiaobo_liu 2008-11-20
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


