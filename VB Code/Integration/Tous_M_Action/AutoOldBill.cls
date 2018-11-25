VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoOldBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_oParse As CParse
Private strLanguage As String

Private m_Bill As Object
Private m_objSave As Object
Private m_objItemCheck As Object
Private m_dictItemCache As Object
Private m_dictStockCache As Object

Private m_sDsn As String

Public Enum CheckType
    Check = 0
    UnCheck = 1
End Enum

Public Enum BillType
    Blue = 0
    Red = 1
End Enum

Public Property Let sDsn(ByVal sDsn As String)
    m_sDsn = sDsn
End Property
Public Property Get sDsn() As String
    sDsn = m_sDsn
End Property

Public Function NewBill(ByVal TranType As Long, Optional ByVal lType As BillType = Blue) As Boolean
    Set m_Bill = New CreateOldBill
    NewBill = m_Bill.NewBill(m_sDsn, m_objSave, m_objItemCheck, TranType, lType)
End Function

Public Function SetHeadValue(ByVal FieldName As String, ByVal Value As Variant) As Boolean
    SetHeadValue = m_Bill.SetHeadValue(m_sDsn, m_objSave, m_objItemCheck, FieldName, Value)
End Function

Public Function SetEntryValue(ByVal FieldName As String, ByVal Value As Variant, ByVal ROW As Long) As Boolean
    SetEntryValue = m_Bill.SetEntryValue(m_sDsn, m_objSave, m_objItemCheck, FieldName, Value, ROW)
End Function

Public Function AddRow(ByVal ROW As Long)
    m_Bill.AddRow m_objSave, ROW
End Function

Public Function SaveBill(ByRef InterID As Long, ByRef BillNo As String, ByRef Result As String) As Boolean
    SaveBill = m_Bill.SaveBill(m_sDsn, m_objSave, m_dictItemCache, m_dictStockCache, Result, BillNo, InterID)
End Function

Public Function CheckBill(ByVal TranType As Long, ByVal InterID As Long, ByRef Result As String, Optional ByVal lType As CheckType = Check) As Boolean
    CheckBill = checkBillData(m_sDsn, InterID, TranType, lType, Result)
End Function

Public Function DeleteBill(ByVal TranType As Long, ByVal InterID As Long, ByRef Result As String) As Boolean
    DeleteBill = deleteBillData(m_sDsn, InterID, TranType, Result)
End Function

'--------------------------------------------------

Private Function InitDataEnv(ByVal sToParse As String) As Boolean
    Set m_oParse = New CParse
    If m_oParse.ParseString(sToParse) Then
        Set datasource = New CDataSource
        Set datasource.ParseObject = m_oParse
    Else
        Err.Raise EBS_E_TypeMismatch, "ParseString"
    End If
    strLanguage = datasource.ParseObject.LanguageID
End Function

Private Function checkBillData(ByVal sDsn As String, _
                                ByVal lBillInterID As Long, _
                                ByVal lTranstype As Long, _
                                ByVal lCheck As Long, _
                                ByRef sRetMsg As String) As Boolean
On Error GoTo HERROR
    Dim strBaseInfo As String
    Dim oCheckBill As Object
    Dim strSQL As String
    Dim lRet As Long
    
    Dim rs As ADODB.Recordset
    Dim sBillNo As String
    Dim sErrorInfo As String, lReturnCode As Long, lReCheck As Long, lReCheck2 As Long
    Dim vectCheckItemInfo As KFO.Vector, sErrorInfo2 As String
    Dim dctPara As KFO.Dictionary
    Dim lUserID As Long
    
    Dim dctTempalteData As KFO.Dictionary
    Dim cn As ADODB.Connection
    
    InitDataEnv sDsn
    Set cn = datasource.Connection
        Dim x As CParse
        Set x = New CParse
        x.ParseString sDsn
        lUserID = x.UserID
        strBaseInfo = x.ConStr & "|" & x.UserID & "|" & x.UserName
        Set x = Nothing
        
        Set dctTempalteData = GetBillTemplateInfo(cn, lTranstype)
            
        If Not (dctTempalteData Is Nothing) Then
            If Len(dctTempalteData.GetValue("HeadTable", "")) > 0 _
                And Len(dctTempalteData.GetValue("BillTemplateID", "")) > 0 _
                And Len(dctTempalteData.GetValue("BillNoField", "")) > 0 _
                And Len(dctTempalteData.GetValue("BillerField", "")) > 0 Then
                
                sBillNo = ""
                strSQL = "SELECT " & dctTempalteData.Value("BillNoField") & " FROM " & dctTempalteData.Value("HeadTable") & " WHERE FTranType=" & lTranstype & " AND FInterID=" & lBillInterID
                Set rs = cn.Execute(strSQL)
                If rs.RecordCount > 0 Then
                    sBillNo = CNulls(rs.Fields(dctTempalteData.Value("BillNoField")).Value, "")
                End If
                Set rs = Nothing
                
                If Len(sBillNo) > 0 Then
                    Set dctPara = New KFO.Dictionary
                    dctPara.Value("PropString") = sDsn
                    dctPara.Value("TransType") = lTranstype
                    dctPara.Value("InterID") = lBillInterID
                    dctPara.Value("CheckerID") = lUserID
                    dctPara.Value("CheckSwitch") = lCheck
                    If lCheck = 0 Then
                        dctPara.Value("OperateCode") = 1
                    Else
                        dctPara.Value("OperateCode") = 2
                    End If
                    dctPara.Value("ReCheck") = 1
                    dctPara.Value("Operatetype") = 0
                    dctPara.Value("CheckDate") = VBA.Date
                    dctPara.Value("ReturnCode") = 0
                    dctPara.Value("ReturnString") = ""
                    Set vectCheckItemInfo = New KFO.Vector
                    Set dctPara.Value("vectItemInfo") = vectCheckItemInfo
                    Set vectCheckItemInfo = Nothing
                    dctPara.Value("MultiCheckLevel") = 0
                    dctPara.Value("WorkFlowFlag") = 0
                    dctPara.Value("BillNo") = sBillNo
                            
                    Set oCheckBill = CreateObject("K3MCheckBill.CheckNow")
                    lRet = oCheckBill.CheckBill(dctPara)
                    Set oCheckBill = Nothing
                    
                    lReturnCode = dctPara.GetValue("ReturnCode", 0)
                    sErrorInfo2 = dctPara.GetValue("ReturnString", "")
                    
                    If lReturnCode = 0 Or lReturnCode = 3 Then
                        checkBillData = True
                    Else
                        checkBillData = False
                        sRetMsg = sErrorInfo2 & "(RetCode:" & lReturnCode & ")"
                    End If
                Else
                    checkBillData = False
                    sRetMsg = "Bill No is Empty!"
                End If
            Else
                checkBillData = False
                sRetMsg = "BillNo Field or Biller Field are not exists!"
            End If
        Else
            checkBillData = False
            sRetMsg = "Template data is not exists!"
        End If
    Exit Function
HERROR:
    checkBillData = False
    sRetMsg = "CheckBill Error:" & Err.Description & "(Source:" & Err.Source & ")"
End Function

Private Function GetBillTemplateInfo(ByVal cn As ADODB.Connection, ByVal lTranType As Long) As KFO.Dictionary
                                        
    Dim strSQL As String
    Dim sBillTemplateID As String
    Dim sHeadTable As String
    Dim sBillNoField As String, sBillerField As String
    Dim i As Long
    Dim rs As ADODB.Recordset
    Dim dctData As KFO.Dictionary
    
    strSQL = "SELECT FHeadTable,FTemplateID FROM ICTransactionType WHERE FID= " & lTranType
    Set rs = cn.Execute(strSQL)
    If rs.RecordCount > 0 Then
        sHeadTable = CNulls(rs.Fields("FHeadTable").Value, "")
        sBillTemplateID = CNulls(rs.Fields("FTemplateID").Value, "")
    End If
    Set rs = Nothing
    
    If Len(sHeadTable) > 0 And Len(sBillTemplateID) > 0 Then
        strSQL = "SELECT FCtlType,FFieldName FROM ICTemplate WHERE FID='" & sBillTemplateID & "'"
        Set rs = cn.Execute(strSQL)
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                Select Case CNulls(rs.Fields("FCtlType").Value, 0)
                    Case 4  '单据编号
                        sBillNoField = CNulls(rs.Fields("FFieldName").Value, "")
                    Case 9  '制单人
                        sBillerField = CNulls(rs.Fields("FFieldName").Value, "")
                End Select
                rs.MoveNext
            Next i
        End If
        Set rs = Nothing
    End If
    
    Set dctData = New KFO.Dictionary
    dctData.Value("HeadTable") = sHeadTable
    dctData.Value("BillTemplateID") = sBillTemplateID
    dctData.Value("BillNoField") = sBillNoField
    dctData.Value("BillerField") = sBillerField
    Set GetBillTemplateInfo = dctData
    Set dctData = Nothing
End Function


'/ 单据删除处理，此次删除接口，会导致HookInvoke的二次开发调用
Private Function deleteBillData(ByVal sDsn As String, _
                                ByVal lBillInterID As Long, _
                                ByVal lTranstype As Long, _
                                ByRef sRetMsg As String) As Boolean
On Error GoTo HERROR
    Dim myCParse As CParse
    Dim strBaseInfo As String
    Dim oDeleteBill As Object
    Dim strSQL As String
    Dim bRet As Boolean
    
    Dim dctTempalteData As KFO.Dictionary
    
    Dim rs As ADODB.Recordset
    Dim lListID As Long, lBillCls As Long, lListTemplateID As Long
    Dim sBillNo As String
    Dim sErrorInfo As String, lReturnCode As Long, lReCheck As Long, lReCheck2 As Long
    Dim vectCheckItemInfo As KFO.Vector, sErrorInfo2 As String
    
    Dim cn As ADODB.Connection
    
    InitDataEnv sDsn
    Set cn = datasource.Connection
    
    If IsControlTranType(lTranstype) Then
        Set myCParse = New CParse
        myCParse.ParseString sDsn
        strBaseInfo = myCParse.ConStr & "|" & myCParse.UserID & "|" & myCParse.UserName
        
        sRetMsg = ""
        sBillNo = ""
        
        If cn Is Nothing Then
            InitDataEnv sDsn
            Set cn = datasource.Connection
        End If
        
        Set dctTempalteData = GetBillTemplateInfo(cn, lTranstype)
            
        If Not (dctTempalteData Is Nothing) Then
            If Len(dctTempalteData.GetValue("HeadTable", "")) > 0 _
                And Len(dctTempalteData.GetValue("BillTemplateID", "")) > 0 _
                And Len(dctTempalteData.GetValue("BillNoField", "")) > 0 _
                And Len(dctTempalteData.GetValue("BillerField", "")) > 0 Then
                
                strSQL = "SELECT " & dctTempalteData.GetValue("BillNoField", "") & " FROM " & dctTempalteData.GetValue("HeadTable", "") & " WHERE FTranType=" & lTranstype & " AND FInterID=" & lBillInterID
                Set rs = cn.Execute(strSQL)
                If rs.RecordCount > 0 Then
                    sBillNo = CNulls(rs.Fields(dctTempalteData.GetValue("BillNoField", "")).Value, "")
                End If
                Set rs = Nothing
            End If
        End If
        
        Select Case lTranstype
            Case 71     '采购订单
                lListID = 26
                lBillCls = 101
                lListTemplateID = 81
            Case 1      '外构入库
                lListID = 1
                lBillCls = 200
                lListTemplateID = 58
            Case 2      '产品入库
                lListID = 2
                lBillCls = 200
                lListTemplateID = 55
            Case 10     '其他入库
                lListID = 7
                lBillCls = 200
                lListTemplateID = 57
            Case 81     '销售订单
                lListID = 32
                lBillCls = 300
                lListTemplateID = 61
            Case 21     '销售出库
                lListID = 8
                lBillCls = 201
                lListTemplateID = 76
            Case 24     '领料单
                lListID = 11
                lBillCls = 201
                lListTemplateID = 77
            Case 29     '其他出库单
                lListID = 16
                lBillCls = 201
                lListTemplateID = 68
            Case 41     '调拨单
                lListID = 18
                lBillCls = 202
                lListTemplateID = 75
            Case 40     '盘盈入库单
                lListID = 17
                lBillCls = 204
                lListTemplateID = 50
            Case 43     '盘亏毁损单
                lListID = 19
                lBillCls = 204
                lListTemplateID = 51
            Case 75     '采购发票
                lListID = 29
                lBillCls = 103
                lListTemplateID = 84
            Case 76     '采购发票
                lListID = 30
                lBillCls = 103
                lListTemplateID = 84
            Case 80     '销售发票(专用)
                lListID = 31
                lBillCls = 302
                lListTemplateID = 62
            Case 86     '销售发票(普通)
                lListID = 37
                lBillCls = 302
                lListTemplateID = 62
            Case 6     '虚仓入库单
                lListID = 4
                lBillCls = 250
                lListTemplateID = 54
            Case 26     '虚仓出库单
                lListID = 14
                lBillCls = 260
                lListTemplateID = 79
            Case 92
                lListID = 103
                lBillCls = 114
                lListTemplateID = 92
            Case 137
                lListID = 137
                lBillCls = 137
                lListTemplateID = 137
                
            Case 5  '委外加工入库单 fantom 2010 -8 -11
                lListID = 5
                lBillCls = 200
                lListTemplateID = 56
                
            Case 28 '委外加工出库 fantom 2010 - 08 -14
                lListID = 15
                lBillCls = 201
                lListTemplateID = 69
                
            
        End Select
        
        If Len(sBillNo) > 0 Then
            Set oDeleteBill = CreateObject("K3ListServer.clsListUpdate")
            lReCheck = 1
            bRet = oDeleteBill.DeleteBill(sBillNo, _
                                            sDsn, _
                                            lListID, _
                                            strBaseInfo, _
                                            lBillInterID, _
                                            lTranstype, _
                                            lBillCls, _
                                            lListTemplateID, _
                                            sErrorInfo, _
                                            0, _
                                            lReturnCode, _
                                            lReCheck, _
                                            lReCheck2, _
                                            vectCheckItemInfo, _
                                            sErrorInfo2)
            Set oDeleteBill = Nothing
            If bRet = True Then
                deleteBillData = True
                sRetMsg = ""
            Else
                deleteBillData = False
                sRetMsg = sErrorInfo & "(RetCode:" & lReturnCode & ")" & sErrorInfo2
            End If
        Else
'            deleteBillData = False
'            sRetMsg = "Bill No is Empty!"
            '没有对应的单号，也不处理 fantom 2011-05-20
            deleteBillData = True
            sRetMsg = ""
        End If
    Else
        deleteBillData = True
        sRetMsg = ""
    End If
    Exit Function
HERROR:
    deleteBillData = False
    sRetMsg = Err.Description & "(Source:" & Err.Source & ")"
End Function

Private Function IsControlTranType(ByVal lTranstype As Long) As Boolean
    IsControlTranType = True
    Select Case lTranstype
        Case 71     '采购订单
        Case 1      '外构入库
        Case 2      '产品入库
        Case 10     '其他入库
        Case 81     '销售订单
        Case 21     '销售出库
        Case 24     '领料单
        Case 29     '其他出库单
        Case 41     '调拨单
        Case 40     '盘盈入库单
        Case 43     '盘亏毁损单
        Case 75     '采购发票
        Case 76     '采购发票
        Case 80     '销售发票(专用)
        Case 86     '销售发票(普通)
        Case 6     '虚仓入库单
        Case 26     '虚仓出库单
        Case 92
        Case 137
        Case 5    '委外加工入库单 fantom 2010 -8 -11
        Case 28   '委外加工出库单 fantom 2010 -8 -14
        Case Else
            IsControlTranType = False
    End Select
End Function

Private Sub Class_Terminate()
'    Set m_Bill = Nothing
'    Set m_objSave = Nothing
'    Set m_objItemCheck = Nothing
'    Set m_dictItemCache = Nothing
'    Set m_dictStockCache = Nothing
End Sub




