VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreateOldBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const C_CHECKBILL = "K3DefineBill.BillTemplateInfo"
Private Const C_BILLDATAACCESS = "BillDataAccess.GetData"
'Private Const C_BILLPACKAGE = "K3Bills.clsBillPackage"
'新加的中间层组件
Private Const C_MBILLPACKAGE = "BillDataAccess.clsBillPackage"


Public Function NewBill(ByVal sDsn As String, ByRef objBillSave As Object _
                        , ByRef objCheckItem As Object, ByVal lTranType As Long, Optional ByVal lRed As Long = 0) As Boolean
On Error GoTo HERROR
    Dim lRet As Long
    Dim lUserID As Long
    Dim sUserName As String
    Dim lROB As Long
    NewBill = False

'    m_lRow = 0
    Set objBillSave = Nothing

    'Set objBillSave = GetObjectContext.CreateInstance("BillDataAccess.clsBillPackage")  '生成个打包对象....要换成中间层的
    Set objBillSave = CreateObject("BillDataAccess.clsBillPackage")  '生成个打包对象....要换成中间层的

    lRet = objBillSave.GetBillPackage(sDsn, lTranType, 0)

    If lRet = 1 Then
        lROB = objBillSave.BillInfo.GetValue("FROB", 0)
        
        Dim cp As CParse
        Set cp = New CParse
        cp.ParseString sDsn
        
        sUserName = cp.UserID
        
        If Len(sUserName) > 0 Then
            SetHeadValue sDsn, objBillSave, objCheckItem, "FBillerID", Val(sUserName)
        End If

        If lROB = 1 Then
            If lRed = 0 Then
                SetHeadValue sDsn, objBillSave, objCheckItem, "FROB", 1
            Else
                SetHeadValue sDsn, objBillSave, objCheckItem, "FROB", -1
            End If
        End If
        NewBill = True
    Else
        NewBill = False
        Set objBillSave = Nothing
    End If
    Exit Function
HERROR:
    NewBill = False
    Set objBillSave = Nothing
    Err.Raise Err.Number, "NewBill Error:" & Err.Source, Err.Description
End Function

Private Function IsReady(ByRef objBillSave As Object) As Boolean
    If Not (objBillSave Is Nothing) Then
        IsReady = True
    Else
        IsReady = False
    End If
End Function

Public Function SetHeadValue(ByVal sDsn As String, ByRef objBillSave As Object, ByRef objCheckItem As Object _
                            , ByVal FieldName As String, ByVal vValue As Variant) As Boolean
On Error GoTo HERROR
    If IsReady(objBillSave) = True Then
        If DoCheckItem(sDsn, objBillSave, objCheckItem, FieldName, vValue, True) = True Then
'            m_oBillSave.SetValue FieldName, vValue, True, 0, 0
            objBillSave.setValue FieldName, vValue, True, 0, 0
'            Debug.Print objBillSave.GetValue(FieldName)
            '_____update jet
'            m_oBillSave.SetValue FieldName, vValue, True, 0, 0 '此中间层组件也昭旧
            '________end
            SetHeadValue = True
        Else
            SetHeadValue = False
            Err.Raise vbObjectError - 1, , "check field [ " & FieldName & " ] value Fail!"
        End If
    Else
        SetHeadValue = False
        Err.Raise vbObjectError - 2, , "BillSave Init Error!"
    End If
    Exit Function
HERROR:
    SetHeadValue = False
    Err.Raise Err.Number, "SetHeadValue Error:" & Err.Source, Err.Description
End Function

Public Function SetEntryValue(ByVal sDsn As String, ByRef objBillSave As Object, ByRef objCheckItem As Object _
                , ByVal FieldName As String, ByVal vValue As Variant, ByVal Row As Long) As Boolean
On Error GoTo HERROR
    If IsReady(objBillSave) = True Then
        If DoCheckItem(sDsn, objBillSave, objCheckItem, FieldName, vValue, False) = True Then
'            m_oBillSave.SetValue FieldName, vValue, False, Row, 0
            objBillSave.setValue FieldName, vValue, False, Row, 0
            'update jet
            '就用上面这句..........
            'end
            SetEntryValue = True
        Else
            SetEntryValue = False
            Err.Raise vbObjectError - 1, , "check field [ " & FieldName & " ] value Fail!"
        End If
    Else
        SetEntryValue = False
        Err.Raise vbObjectError - 2, , "BillSave Init Error!"
    End If
    Exit Function
HERROR:
    SetEntryValue = False
    Err.Raise Err.Number, "SetEntryValue Error:" & Err.Source, Err.Description
End Function

Public Function AddRow(ByVal objSave As Object, ByVal lRow As Long) As Long
On Error GoTo HERROR
    Dim lRowIndex As Long
    
    If IsReady(objSave) = True Then
        Do Until lRowIndex = lRow
            lRowIndex = lRow
            lRowIndex = objSave.AddRow(lRowIndex)
        Loop
    Else
        AddRow = -1
        Err.Raise vbObjectError - 2, , "BillSave Init Error!"
    End If
    Exit Function
HERROR:
    AddRow = -1
    Err.Raise Err.Number, "AddRow Error:" & Err.Source, Err.Description
    
End Function

Public Function SaveBill(ByVal sDsn As String, ByRef objBillSave As Object, ByRef dictItemCache As KFO.Dictionary _
                    , ByRef dictStockCache As KFO.Dictionary, Optional ByRef sMsg As String, _
                        Optional ByRef sReturnBillNo As String, _
                        Optional ByRef sReturnInterID As Long, _
                        Optional ByRef retObject As Object) As Boolean
On Error GoTo HERROR
    Dim oSaveBill As Object
    Dim bCheck As Boolean
    Dim retMsg As String
    Dim oBillData As Object
    
    Dim dicBillStruct As KFO.Dictionary
    Dim dicBillInfo As KFO.Dictionary
    Dim dicSaveInfo As KFO.Dictionary
    
'edited by fq,如果存在单据内码，即是修改单据。
    If sReturnInterID > 0 Then
        objBillSave.BillDataValue(FID) = sReturnInterID
    End If
        
    If IsReady(objBillSave) = True Then
        retMsg = ""
        bCheck = CheckMustInput(sDsn, objBillSave, retMsg, dictItemCache, dictStockCache)
        
        If Len(retMsg) = 0 Then
            
            bCheck = CheckBillDate(sDsn, objBillSave, retMsg)
            If bCheck = True Then
                
                Set oBillData = CreateObject("BillDataAccess.GetData")
                
                Dim dict As KFO.Dictionary
                Set dict = objBillSave.BillPackage '数据包
                
                If Not dict Is Nothing Then
                    Set dicBillStruct = dict.GetValue("DataStructure", Nothing)
                End If
                
                If Not dicBillStruct Is Nothing Then
                    Set dicBillInfo = dicBillStruct.GetValue("BillInfo", Nothing)
                End If
                
                If Not dicBillInfo Is Nothing Then
                    Set dicSaveInfo = dicBillInfo.GetValue("SaveInfo", Nothing)
                End If
                
                If Not dicSaveInfo Is Nothing Then
                    dicSaveInfo.Value("ReSave") = 1
                End If
                
                Set oSaveBill = oBillData.SaveBill(sDsn, dict) '保存数据包
                
                SaveBill = oSaveBill.GetValue("Success", False)
                
                If SaveBill Then
                    sReturnBillNo = oSaveBill.GetValue("ReturnBillNo", "")
                Else
                    sMsg = oSaveBill.GetValue("ReturnMsg", "")
                    Err.Raise vbObjectError - 1, , sMsg
                End If
                
                sReturnInterID = oSaveBill.GetValue("BillInterID", 0)
                Set retObject = oSaveBill
                Set oSaveBill = Nothing
            Else
                SaveBill = False
                Err.Raise vbObjectError - 1, , retMsg
            End If
        Else
            SaveBill = False
            Err.Raise vbObjectError - 1, , retMsg
        End If
    Else
        Err.Raise vbObjectError - 2, , "BillSave Init Error!"
        SaveBill = False
    End If
    Exit Function
HERROR:
    SaveBill = False
    Err.Raise Err.Number, "SaveBill Error:" & Err.Source, Err.Description
End Function




Private Function DoCheckItem(ByVal sDsn As String, ByRef objBillSave As Object, ByRef objItemCheck As Object, ByVal FieldName As String, _
                                ByVal vValue As String, _
                                ByVal isHead As Boolean) As Boolean
On Error GoTo HERROR
    Dim dctField As KFO.Dictionary
    Dim lInterID As Long
    Dim lCtlType As Long, lLookupCls As Long
    Dim dctRet As KFO.Dictionary
    
    DoCheckItem = False
    If Len(vValue) > 0 Then  ''CheckItem = True And
        If objItemCheck Is Nothing Then
            Set objItemCheck = CreateObject("K3DefineBill.BillTemplateInfo")
        End If
        
        Set dctField = objBillSave.GetField(FieldName, isHead)
        '_____add jet
        '照旧
        '____end jet
        If Not dctField Is Nothing Then
            lCtlType = dctField.GetValue("FCtlType", 0)
            lLookupCls = dctField.GetValue("FLookupCls", 0)
        End If
        Set dctField = Nothing
        
        If IsCheckType(lCtlType, lLookupCls) Then
            lInterID = Val(vValue)
            If lInterID <> 0 Then
            '##################这里须更改.............
                Set dctRet = objItemCheck.GetItemDetail(sDsn, lCtlType, lLookupCls, "", lInterID, False)
                If Not (dctRet Is Nothing) Then
                    DoCheckItem = True
                Else
                    DoCheckItem = False
                End If
                Set dctRet = Nothing
            Else
                DoCheckItem = True
            End If
        Else
            DoCheckItem = True
        End If
    Else
        DoCheckItem = True
    End If
    Exit Function
HERROR:
    DoCheckItem = False
    Err.Raise Err.Number, "DoCheckItem Error:" & Err.Source, Err.Description
End Function

Private Function IsCheckType(ByVal lCtlType As Long, ByVal lLookupCls As Long) As Boolean
    Select Case lCtlType
        Case 2, 8, 18, 9, 25
            If lCtlType = 2 And lLookupCls = 6 Then
                IsCheckType = False
            Else
                IsCheckType = True
            End If
        Case Else
            IsCheckType = False
    End Select
End Function

Private Sub Class_Initialize()
'    m_lRow = 0
'    m_bInit = False
'    m_bBillInit = False
'    m_sErrorMsg = ""
'    m_bCheckItem = False
'    Set m_oBillSave = Nothing
'    Set m_oItemCheck = Nothing
'    Set m_dctItemCache = Nothing
'    Set m_dctStockCache = Nothing
End Sub

Private Sub Class_Terminate()
'    m_lRow = 0
'    m_bInit = False
'    m_bBillInit = False
'    m_sErrorMsg = ""
'    m_bCheckItem = False
'    Set m_oBillSave = Nothing
'    Set m_oItemCheck = Nothing
'    Set m_dctItemCache = Nothing
'    Set m_dctStockCache = Nothing
End Sub

Private Function CheckBillDate(ByVal sDsn As String, ByRef objBillSave As Object, ByRef sMsg As String) As Boolean
    Dim oBillData As Object
    Dim lRet As Long
    Set oBillData = CreateObject(C_BILLDATAACCESS)
    '##################这里须更改.............

    lRet = oBillData.CheckBillPackage(sDsn, objBillSave.BillPackage, sMsg)

    'add jet
    
    'end jet
    Set oBillData = Nothing
    If Len(sMsg) > 0 Then
        CheckBillDate = False
    Else
        CheckBillDate = True
    End If
End Function



Private Function CheckMustInput(ByVal sDsn As String, ByRef objBillSave As Object _
                                , ByRef sMsg As String, ByRef dctItemCache As KFO.Dictionary, ByRef dctStockCache As KFO.Dictionary) As Boolean
    Dim i As Long, lRow As Long
    Dim strStep As String, rs As ADODB.Recordset
    Dim vctHead As KFO.Vector, vctEntry As KFO.Vector
    Dim dctHeadField As KFO.Dictionary, dctHeadValue As KFO.Dictionary
    Dim dctEntryField As KFO.Dictionary, vecEntrys As KFO.Vector
    
    Dim bMustInput As Boolean, bEnable As Boolean, bNeedSave As Boolean
    Dim sCaption As String, lCtlType As Long, sFieldName As String, lLookupCls As Long
    Dim vValue As Variant
    
    Dim sSaveRule As String, strFunc As String
    Dim lTranType As Long
    Dim bSpecialFor6_26 As Boolean
    Dim bGoNext As Boolean
    
    CheckMustInput = False
    bSpecialFor6_26 = False
    
    
    Set vctHead = objBillSave.BillDataHeadStruct
    Set dctHeadValue = objBillSave.BillDataHeadValue
    
    Set vctEntry = objBillSave.BillDataEntryStruct
    Set vecEntrys = objBillSave.BillDataEntrysValue
  
    
    sMsg = ""
    'add by lai 081028
    lTranType = objBillSave.BillTranType
    With objBillSave
    For i = vctHead.LBound To vctHead.UBound
        Set dctHeadField = vctHead.Item(i)
        bMustInput = CBool(dctHeadField.GetValue("FMustInput", False))
        bEnable = GetEnable(dctHeadField.GetValue("FEnable", 0))
        bNeedSave = CBool(dctHeadField.GetValue("FNeedSave", False))
        
        If bMustInput = True And bEnable = True And bNeedSave = True Then
            sCaption = dctHeadField.GetValue("FCaption", "")
            lCtlType = dctHeadField.GetValue("FCtlType", 0)
            sFieldName = dctHeadField.GetValue("FFieldName", "")
            
            If lCtlType <> 4 Then
                vValue = Trim(.GetValue(sFieldName, "", True, 0, 0))
                
                If (lTranType = 6 Or lTranType = 26 Or lTranType = 10 Or lTranType = 29) And (UCase(sFieldName) = UCase("FCustID") Or UCase(sFieldName) = UCase("FDeptID") Or UCase(sFieldName) = UCase("FSupplyID")) Then
                    If Len(vValue) > 0 Then
                        bSpecialFor6_26 = True
                    End If
                    
                Else
                    
                    If Len(vValue) = 0 Then
                        CheckMustInput = True
                        sCaption = Replace(sCaption, ":", "")
                        
                        sMsg = "数据不完整，不能保存，必须输入项 [" & sCaption & "]"
                        
                        
                        GoTo HOver
                    ElseIf IsNumberField(lCtlType) Then
                        If Val(vValue) = 0 Then
                            CheckMustInput = True
                            sCaption = Replace(sCaption, ":", "")
                            sMsg = "数据不完整，不能保存，必须输入项 [" & sCaption & "]"
                            
                            GoTo HOver
                        End If
                    End If
                End If
            End If
        End If
        Set dctHeadField = Nothing
    Next i
    
    If (lTranType = 6 Or lTranType = 26 Or lTranType = 10 Or lTranType = 29) And bSpecialFor6_26 = False Then
        CheckMustInput = True
        sMsg = "部门、供应商、客户至少录入一个！"
        GoTo HOver
    End If
    
    For lRow = 1 To vecEntrys.UBound
        vValue = Trim(.GetValue(objBillSave.BillKeyItemID, "", False, lRow, 0))
        
        
        If vValue <> "" Then
            For i = vctEntry.LBound To vctEntry.UBound
            
                bGoNext = False
                
                Set dctEntryField = vctEntry.Item(i)
                bMustInput = CBool(dctEntryField.GetValue("FMustInput", False))
                bEnable = GetEnable(dctEntryField.GetValue("FEnable", 0))
                bNeedSave = CBool(dctEntryField.GetValue("FNeedSave", False))
                
                sCaption = dctEntryField.GetValue("FHeadCaption", "")
                lCtlType = dctEntryField.GetValue("FCtlType", 0)
                lLookupCls = dctEntryField.GetValue("FLookupcls", 0)
                sFieldName = dctEntryField.GetValue("FFieldName", "")
                sSaveRule = dctEntryField.GetValue("FSaveRule", "")
                
                '针对V12.1生成单据时的特殊处理,只用于特定情况
                '------------------------------------
                If sFieldName = "FMTONo" Then       '
                    bNeedSave = False               '
                End If                              '
                '------------------------------------
                
                bGoNext = IsItemPropCheck(sDsn, dctItemCache, dctStockCache, objBillSave, dctEntryField, lRow)
                
                If bGoNext = True Then
                    If bMustInput = True And bEnable = True And bNeedSave = True Then
                        vValue = Trim(.GetValue(sFieldName, "", False, lRow, 0))
                        If vValue = "" Or vValue = "0" Then
                            CheckMustInput = True
                            sMsg = "数据不完整,不能保存。必须输入第" & lRow & "行的分录项 [" & sCaption & "]"
                            GoTo HOver
                        ElseIf IsNumberField(lCtlType) Then
                            If Val(vValue) = 0 Then
                                CheckMustInput = True
                                sMsg = "数据不完整,不能保存。必须输入第" & lRow & "行的分录项 [" & sCaption & "]"
                                GoTo HOver
                            End If
                        End If
                    End If
                    
                    If bNeedSave = True Then
                        vValue = Trim(.GetValue(sFieldName, "", False, lRow, 0))
                        If IsDate(lCtlType) = True Then
                            If Len(vValue) > 0 Then
                                If Not (IsDateField(vValue)) Then
                                    CheckMustInput = True
                                    sMsg = "数据不完整,不能保存。必须输入第" & lRow & "行的分录项 [" & sCaption & "]"
                                    GoTo HOver
                                End If
                            End If
                        End If
                        
                        Do While Len(sSaveRule) > 0
                            If InStr(sSaveRule, ";") > 0 Then
                                strStep = VBA.Left(sSaveRule, InStr(sSaveRule, ";") - 1)
                                sSaveRule = Mid(sSaveRule, InStr(sSaveRule, ";") + 1)
                            Else
                                strStep = sSaveRule
                                sSaveRule = ""
                            End If
                            strFunc = ""
                            strFunc = VBA.Left(strStep, 1)
                            strStep = Mid(strStep, 3)
                            If Not SetPropDesc(objBillSave, lRow, strFunc, strStep, dctEntryField, vValue, sMsg) Then
                                CheckMustInput = True
                                GoTo HOver
                            End If
                        Loop
                    End If
                End If
                Set dctEntryField = Nothing
            Next i
            Set rs = Nothing
        End If
    Next lRow
    
    End With
    Exit Function
HOver:
    CheckMustInput = False
    Err.Raise vbObjectError - 1, "CheckMustInput Error:", sMsg
    Exit Function
HERROR:
    CheckMustInput = False
    Err.Raise Err.Number, "CheckMustInput Error:" & Err.Source, Err.Description
End Function

Private Function GetEnable(ByVal EnableValue As Integer) As Boolean
    GetEnable = EnableValue And 32
End Function

Private Function IsNumberField(ByVal lCtlType As Long) As Boolean
    Select Case lCtlType
        Case 3, 11, 12, 13, 20, 21, 31, 33
            IsNumberField = True
        Case Else
            IsNumberField = False
    End Select
End Function

Private Function IsDateField(ByVal lCtlType As Long) As Boolean
    Select Case lCtlType
        Case 1, 17, 32
            IsDateField = True
        Case Else
            IsDateField = False
    End Select
End Function

Private Function SetPropDesc( _
    ByRef objBillSave As Object, _
    ByVal Row As Long, _
    ByVal chrF As String, _
    ByVal strF As String, _
    ByVal dctField As KFO.Dictionary, _
    ByVal fieldValue As Variant, _
    ByRef sMsg As String) As Boolean

On Error GoTo HERROR
    Dim t1, t As Variant
    Dim sSplit As String
    Dim sCaption As String
    SetPropDesc = True
    
    sSplit = " "
    
    sCaption = dctField.GetValue("FHeadCaption", "")
    Select Case chrF
        Case "@"
            If Trim(fieldValue) = "" Then
'                sMsg = sCaption & "不能为空"
'                SetPropDesc = False
'                GoTo HOver
            End If
        Case "U"
            If Val(fieldValue) = Val(strF) Then
                sMsg = "输入的数据: " & sCaption & sSplit & "  不能为" & sSplit & strF
                SetPropDesc = False
                GoTo HOver
            End If
        Case ">"
            If strF = "CurrentDate" Then
            Else
                If IsNumeric(strF) Then
                    t1 = strF
                Else
                    t1 = objBillSave.GetValue(strF, "", False, Row, 0)
                    'add jet
                    
                    'end jet
                End If
                If Val(fieldValue) <= Val(t1) Then
                    If IsNumeric(strF) Then
                        sMsg = strF
                    Else
                        sMsg = " [" & objBillSave.GetField(strF, False).GetValue("FHeadCaption", "") & "] 的值"
                    End If
                    sMsg = "输入的数据:[" & CStr(fieldValue) & sSplit & "] 必须大于" & sSplit & sMsg
                    SetPropDesc = False
                    GoTo HOver
                End If
            End If
        Case "^"
            If Val(fieldValue) = 0 Then
                SetPropDesc = False
                sMsg = "输入的数据:[" & CStr(fieldValue) & sSplit & "] 必须大于0"
                GoTo HOver
            End If
            
            If IsNumeric(strF) Then
                t1 = strF
            Else
                t1 = objBillSave.GetValue(strF, "", False, Row, 0)
            End If
            If Val(fieldValue) < Val(t1) Then
                If IsNumeric(strF) Then
                    sMsg = strF
                Else
                    sMsg = " [" & objBillSave.GetField(strF, False).GetValue("FHeadCaption", "") & "] 的值"
                End If
                sMsg = "输入的数据:[" & CStr(fieldValue) & sSplit & "] 必须大于或等于" & sSplit & sMsg
                SetPropDesc = False
                GoTo HOver
            End If
        Case "~"
            If Val(fieldValue) < 0 Then
                SetPropDesc = False
                sMsg = "输入的数据:[" & CStr(fieldValue) & sSplit & "] 不能小于0"
                GoTo HOver
            End If
            
            If Trim(strF) = "" Then Exit Function
            If IsNumeric(strF) Then
                t1 = strF
            Else
                t1 = objBillSave.GetValue(strF, "", False, Row, 0)
            End If
            
            If Val(fieldValue) >= Val(t1) Then
                If IsNumeric(strF) Then
                    sMsg = strF
                Else
                    sMsg = " [" & objBillSave.GetField(strF, False).GetValue("FHeadCaption", "") & "] 的值"
                End If
                sMsg = "输入的数据:[" & CStr(fieldValue) & sSplit & "] 必须小于" & sSplit & sMsg
                SetPropDesc = False
                GoTo HOver
            End If
        Case "#"
'            If Val(fieldValue) <= 0 Then
'                sMsg = "输入的数据: " & sCaption & sSplit & " 必须大于0"
'                SetPropDesc = False
'                GoTo HOver
'            End If
        Case "V"   '分录日期大于单据头日期
            If Trim(fieldValue) = "" Then Exit Function
            t1 = GetString(strF, 1)
            If t1 <> "" Then t1 = objBillSave.GetValue(t1, "", True, 0, 0)
            If t1 <> "" And CDate(t1) > CDate(fieldValue) Then
                sMsg = "第" & Row & "条分录的输入项[" & sCaption & "] 不应小于 [" & objBillSave.GetField(GetString(strF, 1), True).GetValue("FCaption", "") & "]"
                SetPropDesc = False
                GoTo HOver
            End If
        Case "$"
            If Val(fieldValue) <= 0 Then
                t1 = objBillSave.GetValue(strF, "", False, Row, 0)
                If Val(fieldValue) = 0 And Val(t1) = 0 Then
                    sMsg = sCaption & " 与 " & objBillSave.GetField(strF, False).GetValue("FHeadCaption", "") & " 的值不能同时为0"
                    SetPropDesc = False
                    GoTo HOver
                End If
                If Val(fieldValue) < 0 And Val(t1) > 0 Then
                    sMsg = " 不允许存在" & sSplit & sCaption & sSplit & "(值:" & fieldValue & ")为负、" & objBillSave.GetField(strF, False).GetValue("FHeadCaption", "") & "(值:" & t1 & ")为正的分录"
                    SetPropDesc = False
                    GoTo HOver
                End If
            End If
         Case "D"
            If IsDate(strF) Then
                t1 = strF
            Else
                t1 = objBillSave.GetValue(strF, "", False, Row, 0)
            End If
            If Trim(t1) <> "" And fieldValue < t1 Then
                If IsDate(strF) Then
                    sMsg = strF
                Else
                    sMsg = " [" & objBillSave.GetField(strF, False).GetValue("FHeadCaption", "") & "] 的值"
                End If
                sMsg = "输入的数据:[" & sSplit & CStr(sCaption) & sSplit & "] 必须大于或等于" & sSplit & sMsg
                SetPropDesc = False
                GoTo HOver
            End If
        Case Else
    End Select
    
    Exit Function
HOver:
    Exit Function
HERROR:
    SetPropDesc = False
    Debug.Assert False
End Function

Private Function GetString( _
    ByVal str As String, _
    Optional ByVal place As Integer = 1, _
    Optional ByVal sep As String = ",") As String
    
    Dim sTmp As String
    Dim sReturn As String
    Dim i As Integer
    Dim l As Integer
    On Error Resume Next
    
    sTmp = str & sep
    For l = 1 To place
        i = InStr(sTmp, sep)
        If i = 0 Then
            sReturn = ""
            GoTo ExitSub
        End If
        sReturn = VBA.Left(sTmp, i - 1)
        sTmp = Right(sTmp, Len(sTmp) - i)
    Next
ExitSub:
    GetString = sReturn
End Function

Private Function GetItemProp(ByVal sDsn As String, ByRef dctItemCache As KFO.Dictionary, ByRef objBillSave As Object _
                        , ByVal lRow As Long, ByVal valueType As Long) As Long
On Error GoTo HERROR
    Dim oBillData As Object
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim vValue As Variant
    Dim lItemID As Long
    Dim lRet As Long
    Dim lPlanMode As Long
    
    vValue = Trim(objBillSave.GetValue(objBillSave.BillKeyItemID, "", False, lRow, 0))
    
    If Len(vValue) > 0 Then
    
        lItemID = Val(vValue)
        
        If dctItemCache Is Nothing Then
            Set dctItemCache = New KFO.Dictionary
        End If
        
        If dctItemCache.Lookup(lItemID) = True Then
            Set rs = dctItemCache.GetValue(lItemID, Nothing)
        End If
        
        If rs Is Nothing Then
            strSQL = "SELECT * From t_ICItem WHERE FItemID=" & lItemID
            Set oBillData = CreateObject(C_BILLDATAACCESS)
            '##################这里须更改.............
            Set rs = oBillData.ExecuteSQL(sDsn, strSQL)
            Set oBillData = Nothing
            Set dctItemCache.Value(lItemID) = rs
        End If
                
        If Not (rs Is Nothing) Then
            If rs.RecordCount > 0 Then
                Select Case valueType
                    Case 1
                        lRet = CNulls(rs.Fields("FAuxClassID").Value, 0)
                    Case 2
                        lRet = CNulls(rs.Fields("FBatchManager").Value, 0)
                    Case 3
                        lRet = CNulls(rs.Fields("FISKFPeriod").Value, 0)
                    Case 4
                        lRet = CNulls(rs.Fields("FErpClsID").Value, 0)
                    Case 5
                        On Error Resume Next
                        lPlanMode = CNulls(rs.Fields("FPlanMode").Value, 0)
                        On Error GoTo 0
                        If lPlanMode <> 14035 Then
                            lRet = 0
                        Else
                            lRet = 1
                        End If
                    Case Else
                        lRet = 0
                End Select
            End If
        End If
        Set rs = Nothing
        
        GetItemProp = lRet
    End If
    Exit Function
HERROR:
    Set rs = Nothing
    GetItemProp = False
End Function

Private Function GetStockProp(ByVal sDsn As String, ByRef dctStockCache As KFO.Dictionary, ByRef objBillSave As Object _
            , ByVal FieldName As String, ByVal lRow As Long, ByVal valueType As Long) As Boolean
On Error GoTo HERROR
    Dim oBillData As Object
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim vValue As Variant
    Dim lItemID As Long
    Dim lRet As Long
    vValue = Trim(objBillSave.GetValue(FieldName, "", False, lRow, 0))
    
    If Len(vValue) > 0 Then
    
        lItemID = Val(vValue)
    
        If dctStockCache Is Nothing Then
            Set dctStockCache = New KFO.Dictionary
        End If
        
        If dctStockCache.Lookup(lItemID) = True Then
            Set rs = dctStockCache.GetValue(lItemID, Nothing)
        End If
        
        If rs Is Nothing Then
            strSQL = "SELECT FIsStockMgr FROM t_Stock WHERE FItemID=" & lItemID
            Set oBillData = CreateObject(C_BILLDATAACCESS)
            '##################这里须更改.............
            Set rs = oBillData.ExecuteSQL(sDsn, strSQL)
            Set oBillData = Nothing
            Set dctStockCache.Value(lItemID) = rs
        End If
        
        If Not (rs Is Nothing) Then
            If rs.RecordCount > 0 Then
                Select Case valueType
                    Case 1
                        lRet = CNulls(rs.Fields("FIsStockMgr").Value, 0)
                    Case Else
                        lRet = 0
                End Select
            End If
        End If
        Set rs = Nothing
        
        If lRet = 0 Then
            GetStockProp = False
        Else
            GetStockProp = True
        End If
        
    End If
    Exit Function
HERROR:
    Set rs = Nothing
    GetStockProp = False
End Function

Private Function IsItemPropCheck(ByVal sDsn As String, ByRef dctItemCache As KFO.Dictionary, ByRef dctStockCache As KFO.Dictionary, ByRef objBillSave As Object _
                            , ByVal dctEntryField As KFO.Dictionary, ByVal lRow As Long) As Boolean
On Error GoTo HERROR

    Dim lCtlType As Long, sFieldName As String, lLookupCls As Long, sRelateID As String
    Dim bCheck As Boolean
    Dim lRet As Long
    
    Dim vRelationID As Variant
    Dim dctField As KFO.Dictionary
    Dim i As Long
    Dim sRelateFieldName As String
    
    lCtlType = dctEntryField.GetValue("FCtlType", 0)
    lLookupCls = dctEntryField.GetValue("FLookupcls", 0)
    sFieldName = dctEntryField.GetValue("FFieldName", "")
    sRelateID = dctEntryField.GetValue("FRelationID", "")
    
    If lCtlType = 2 And lLookupCls = -99 Then
        lRet = GetItemProp(sDsn, dctItemCache, objBillSave, lRow, 1)
        If lRet = 0 Then
            bCheck = False
            objBillSave.setValue sFieldName, 0, False, lRow, 0
        Else
            bCheck = True
        End If
    ElseIf lCtlType = 10 And lLookupCls = 18 Then
        lRet = GetItemProp(sDsn, dctItemCache, objBillSave, lRow, 2)
        If lRet = 0 Then
            bCheck = False
            objBillSave.setValue sFieldName, "", False, lRow, 0
        Else
            bCheck = True
        End If
    ElseIf UCase(sFieldName) = UCase("FKFDate") Or UCase(sFieldName) = UCase("FKFPeriod") Or UCase(sFieldName) = UCase("FPeriodDate") Then
        lRet = GetItemProp(sDsn, dctItemCache, objBillSave, lRow, 3)
        If lRet = 0 Then
            bCheck = False
            objBillSave.setValue sFieldName, "", False, lRow, 0
        Else
            bCheck = True
        End If
    ElseIf lCtlType = 2 And lLookupCls = -15 Then
        sRelateFieldName = ""
        If Len(sRelateID) > 0 Then
            vRelationID = Split(sRelateID, ",")
            For i = LBound(vRelationID) To UBound(vRelationID)
                Set dctField = objBillSave.GetField(vRelationID(i), False)
                If Not (dctField Is Nothing) Then
                    If dctField.GetValue("FCtlType", 0) = 2 And dctField.GetValue("FLookupcls", 0) = 5 Then
                        sRelateFieldName = dctField.GetValue("FFieldName", "")
                        Exit For
                    End If
                End If
                Set dctField = Nothing
            Next i
        End If
        If Len(sRelateFieldName) > 0 Then
            bCheck = GetStockProp(sDsn, dctStockCache, objBillSave, sRelateFieldName, lRow, 1)
        End If
        If bCheck = False Then
            objBillSave.setValue sFieldName, "", False, lRow, 0
        End If
    ElseIf lCtlType = 2 And lLookupCls = -101 Then
        lRet = GetItemProp(sDsn, dctItemCache, objBillSave, lRow, 4)
        If lRet = 7 Then
            bCheck = True
        Else
            bCheck = False
            objBillSave.setValue sFieldName, "", False, lRow, 0
        End If
    ElseIf UCase(sFieldName) = UCase("FMTONo") Then
        lRet = GetItemProp(sDsn, dctItemCache, objBillSave, lRow, 5)
        If lRet = 0 Then
            bCheck = False
            objBillSave.setValue sFieldName, "", False, lRow, 0
        Else
            bCheck = True
        End If
    Else
        bCheck = True
    End If
        
    
    IsItemPropCheck = bCheck

    Exit Function
HERROR:
    IsItemPropCheck = True
End Function


Private Function GetUserID(ByVal sDsn As String) As String
    Dim strProps As String
    Dim i As Long
    Dim vUserID
    Dim vValue
    strProps = sDsn
    i = InStr(1, strProps, "UserID=", vbTextCompare)
    If i > 0 Then
        strProps = Right(strProps, Len(strProps) - i + 1)
        vUserID = Split(strProps, ";")
        vValue = Right(vUserID(0), Len(vUserID(0)) - Len("UserID="))
        GetUserID = vValue
    End If
    
End Function






