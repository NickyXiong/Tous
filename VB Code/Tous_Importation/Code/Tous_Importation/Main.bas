Attribute VB_Name = "Main"



'---------------------------------------------------------------------------------------
' Procedure : GetDCSPID
' DateTime  : 2013-1-29 15:20
' Author    :
' Purpose   : 查找仓位信息
'---------------------------------------------------------------------------------------
'
Public Function GetDCSPID(ByVal CNN As ADODB.Connection, ByVal sNumber As String, ByVal lSPGroupID As Long) As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "SELECT FSPID FROM t_StockPlace  where FSPGroupID =" & lSPGroupID & " and FNumber='" & Trim(sNumber) & "'"
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetDCSPID = CNulls(rsTemp.Fields("FSPID"), 0)
        End If
    End If
    
    Set rsTemp = Nothing
End Function


''获取仓库组的默认仓位
Public Function GetDEFDCSPID(ByVal CNN As ADODB.Connection, ByVal lSPGroupID As Long) As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "Select FDefaultSPID from t_StockPlaceGroup where FSPGroupID != 0 And FSPGroupID =" & lSPGroupID
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetDEFDCSPID = CNulls(rsTemp.Fields("FDefaultSPID"), 0)
        End If
    End If
    
    Set rsTemp = Nothing
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetIsDCSP
' DateTime  : 2013-1-29 15:35
' Author    :
' Purpose   : 检查是否进行仓位管理
'---------------------------------------------------------------------------------------
'
Public Function GetIsDCSP(ByVal CNN As ADODB.Connection, ByVal lFItemID As Long, ByRef lSPGroupID As Long) As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select FSPGroupID,FIsStockMgr  from t_stock where FItemID=" & lFItemID
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            If CNulls(rsTemp.Fields("FIsStockMgr"), 0) = 1 Or CNulls(rsTemp.Fields("FIsStockMgr"), 0) = True Then
                GetIsDCSP = True
                lSPGroupID = CNulls(rsTemp.Fields("FSPGroupID"), 0)
            End If
        End If
    End If
    Set rsTemp = Nothing
End Function



'---------------------------------------------------------------------------------------
' Procedure : GetLoactionTypeNum
' DateTime  : 2013-1-30 00:32
' Author    :
' Purpose   : 查找仓库属性locationtype number 字段值
'---------------------------------------------------------------------------------------
'
Public Function GetLoactionTypeNum(ByVal CNN As ADODB.Connection, ByVal lStockID As Long) As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select t2.FID,t1.FNumber,t1.FItemID from t_stock t1 " & vbCrLf & _
            "inner join t_submessage t2 on t1.FLocalType=t2.FInterID" & vbCrLf & _
            "Where t1.FItemID = " & lStockID
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetLoactionTypeNum = CNulls(rsTemp.Fields("FID"), "")
        End If
    End If
    Set rsTemp = Nothing
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetStockIDByLoaction
' DateTime  : 2013-2-20 11:43
' Author    : Administrator
' Purpose   : 根据locationtype值取中间仓的仓库
'---------------------------------------------------------------------------------------
'
Public Function GetStockIDByLoaction(ByVal CNN As ADODB.Connection, ByVal sNumber As String) As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select t2.FID,t1.FNumber,t1.FItemID from t_stock t1 " & vbCrLf & _
            "inner join t_submessage t2 on t1.FLocalType=t2.FInterID" & vbCrLf & _
            "Where t2.FID = '" & sNumber & "'"
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetStockIDByLoaction = CNulls(rsTemp.Fields("FItemID"), "")
        End If
    End If
    Set rsTemp = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsExitBill
' DateTime  : 2013-1-30 11:11
' Author    :
' Purpose   : 查询单据是否存在
'---------------------------------------------------------------------------------------
'
Public Function IsExitBill(ByVal CNN As ADODB.Connection, ByVal sBillNo As String, ByVal sTable As String, Optional ByVal lFTranType As Long = 0, Optional ByRef lInterID As Long = 0, Optional sFieldName As String = "FBillNo") As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select FInterID  from " & sTable & " where " & sFieldName & "='" & Trim(sBillNo) & "' and FTranType=" & lFTranType
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            IsExitBill = True
            lInterID = CNulls(rsTemp.Fields("FInterID"), 0)
        End If
    End If
    Set rsTemp = Nothing
End Function


Public Function GetExitBill(ByVal CNN As ADODB.Connection, ByVal sBillNo As String) As ADODB.Recordset
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select v1.FInterID,v1.FROB,sum(abs(u1.FConsignAmount)) as FSumAmt from ICStockBill v1" & vbCrLf & _
            "inner join ICStockBillEntry u1 on v1.FInterID =u1.FInterID" & vbCrLf & _
            "where v1.FTranType=21 and v1.FPosNum like '" & sBillNo & "%'" & vbCrLf & _
            "group by v1.FROB,v1.FInterID "
    Set rsTemp = CNN.Execute(strSql)
    
    Set GetExitBill = rsTemp
    Set CNN = Nothing
    Set rsTemp = Nothing
End Function



'---------------------------------------------------------------------------------------
' Procedure : GetItemID
' DateTime  : 2013-1-25 16:50
' Author    :
' Purpose   : 查找基础资料FItemID
'---------------------------------------------------------------------------------------
'
Public Function GetItemID(ByVal CNN As ADODB.Connection, ByVal sNumber As String, ByVal lItemClassID As Long, Optional ByRef lUnitID As Long = 0, Optional ByRef dRate As Double) As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    If lItemClassID = 4 Then
        strSql = "select FUnitID,FItemID,FTaxRate from t_icitem where FNumber='" & Trim(sNumber) & "'"
    Else
        strSql = "select t2.FName_en,t1.FItemID,t1.fnumber,t1.FName from t_item  t1 " & vbCrLf & _
                "inner join t_itemclass t2 on t1.FItemClassID=t2.FItemClassID" & vbCrLf & _
                "where t1.FDetail=1 and  t1.FshortNumber like '%" & sNumber & "' and t1.FItemClassID= " & lItemClassID
    End If
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetItemID = CNulls(rsTemp.Fields("FItemID"), 0)
            If lItemClassID = 4 Then ''如果是物料需查找单位
                lUnitID = CNulls(rsTemp.Fields("FUnitID"), 0)
                dRate = CNulls(rsTemp.Fields("FTaxRate"), 0)
            End If
        End If
    End If
    
    Set rsTemp = Nothing
    Set CNN = Nothing
End Function



'---------------------------------------------------------------------------------------
' Procedure : GetCustomID
' DateTime  : 2013-2-21 00:13
' Author    :
' Purpose   : 根据部门内码取客户
'---------------------------------------------------------------------------------------
'
Public Function GetCustomID(ByVal CNN As ADODB.Connection, ByVal lDeptID As Long) As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select Flandlordid,* from t_Department where FItemID =" & lDeptID
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetCustomID = CNulls(rsTemp.Fields("FLandlordid"), 0)
        End If
    End If
    Set rsTemp = Nothing

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetSaleType
' DateTime  : 2013-1-28 17:04
' Author    :
' Purpose   : 查找客户属性Store Type值
'---------------------------------------------------------------------------------------
'
Public Function GetSaleType(ByVal CNN As ADODB.Connection, ByVal lItemID As Long) As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select isnull(FStoreType,0)FStoreType from t_Organization where FItemID=" & lItemID
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetSaleType = CNulls(rsTemp.Fields("FStoreType"), 0)
        End If
    End If
    Set rsTemp = Nothing
End Function




'/单据审核处理，lCheck=0，表示审核，lCheck=1，表示反审核，lStockType仓库类型 0代表实仓 ，1代表虚仓
Public Function checkBillData(ByVal sDsn As String, _
                                ByVal lBillInterID As Long, _
                                ByVal lTransType As Long, _
                                ByVal lCheck As Long, _
                                ByRef sRetMsg As String, _
                                Optional ByVal lcheckID As Long = 0, _
                                Optional ByVal lStockType As Long = 0) As Boolean
'On Error GoTo HError
    Dim oCheckBill As Object
    Dim strSql As String
    Dim lRet As Long
    
    Dim rs As ADODB.Recordset
    Dim sErrorInfo As String, lReturnCode As Long, lReCheck As Long, lReCheck2 As Long
    Dim vectCheckItemInfo As KFO.Vector, sErrorInfo2 As String
    Dim dctPara As KFO.Dictionary

    Set dctPara = New KFO.Dictionary
    dctPara.Value("PropString") = sDsn
    dctPara.Value("TransType") = lTransType
    dctPara.Value("InterID") = lBillInterID
    dctPara.Value("CheckerID") = lcheckID 'IIf(lProjectID = 0, 0, getDefaultUserID(sDsn, lTranstype, lProjectID))
    dctPara.Value("CheckSwitch") = lCheck
    If lCheck = 0 Then
        dctPara.Value("OperateCode") = 1
    Else
        dctPara.Value("OperateCode") = 2
    End If
    '增加是否允许实仓负库存判断
'    If lStockType = 0 Then
'        StrSql = "select 1 from t_systemprofile where fkey ='UnderStock' and fvalue=1"
'        Set rs = ExecSQL(sDsn, StrSql)
'        If rs.RecordCount > 0 Then
'            lReCheck = 1
'        Else
'            lReCheck = 0
'        End If
'    Else
'        StrSql = "select 1 from t_systemprofile where fkey ='UnderStockVirtual' and fvalue=1"
'        Set rs = ExecSQL(sDsn, StrSql)
'        If rs.RecordCount > 0 Then
'            lReCheck = 1
'        Else
'            lReCheck = 0
'        End If
'    End If
    dctPara.Value("ReCheck") = lReCheck
    dctPara.Value("Operatetype") = 0
    dctPara.Value("CheckDate") = Date
    dctPara.Value("ReturnCode") = 0
    dctPara.Value("ReturnString") = ""
    Set vectCheckItemInfo = New KFO.Vector
    Set dctPara.Value("vectItemInfo") = vectCheckItemInfo
    Set vectCheckItemInfo = Nothing
    dctPara.Value("MultiCheckLevel") = 0
    dctPara.Value("WorkFlowFlag") = 0
            
    Set oCheckBill = GetObjectContext.CreateInstance("K3MCheckBill.CheckNow")
    lRet = oCheckBill.CheckBill(dctPara)
    Set oCheckBill = Nothing
    
    lReturnCode = dctPara.GetValue("ReturnCode", 0)
    sErrorInfo2 = dctPara.GetValue("ReturnString", "")
    If lReturnCode = 0 Then
        checkBillData = True
    Else
        checkBillData = False
        sRetMsg = sErrorInfo2 & "(RetCode:" & lReturnCode & ")"
        Err.Raise -1, , sRetMsg
    End If
    Exit Function

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetLotInfor
' DateTime  : 2013-1-25 16:50
' Author    :
' Purpose   : 查找基础资料FItemID
'---------------------------------------------------------------------------------------
'
Public Function GetLotInfor(ByVal lItemID As Long, ByVal lStockID As Long, ByVal lSPID As Long, ByRef strLotNo As String, ByRef strKFDate As String, ByRef strKFPeriod As String) As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo Err
    
    '判断是否有启用批次管理或者保质期管理，如果未启用，则直接返回空值
    strSql = "select FBatchManager,FISKFPeriod from t_ICItem where FItemID=" & CStr(lItemID)
    Set rsTemp = modPub.ExecSql(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            If rsTemp.Fields("FBatchManager").Value <> True And rsTemp.Fields("FISKFPeriod").Value <> True Then
                strLotNo = ""
                strKFDate = ""
                strKFPeriod = ""
                GetLotInfor = True
                Set rsTemp = Nothing
                Exit Function
            End If
        End If
    End If
    
    strSql = "select top 1 FBatchNo,FKFDate,FKFPeriod,FQty from ICInventory  "
    strSql = strSql & vbCrLf & "Where FQty<>0 and FItemID=" & CStr(lItemID) & " and FStockID=" & CStr(lStockID) & " and FStockPlaceID=" & CStr(lSPID)
    strSql = strSql & vbCrLf & "group by FBatchNo,FKFDate,FKFPeriod,FQty"
    strSql = strSql & vbCrLf & "order by FKFDate"
    Set rsTemp = modPub.ExecSql(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            strLotNo = Trim(rsTemp.Fields("FBatchNo").Value)
            strKFDate = Trim(rsTemp.Fields("FKFDate").Value)
            strKFPeriod = Trim(rsTemp.Fields("FKFPeriod").Value)
            GetLotInfor = True
        Else
            GetLotInfor = False
        End If
    End If
    
    Set rsTemp = Nothing
    Exit Function
Err:
    GetLotInfor = False
    Set rsTemp = Nothing
    Exit Function
End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckFieldName
' DateTime  : 2014-3-16 16:50
' Author    :
' Purpose   : 查找单据字段名
'---------------------------------------------------------------------------------------
Public Function CheckFieldName(ByVal CNN As ADODB.Connection, strCaption As String, strType As String, iFlag As Integer) As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    'iFlag=0:表头； iFlag=1:表体
    If iFlag = 0 Then
        strSql = "select FFieldName from ICTemplate where FID='" & strType & "' and FCaption='" & strCaption & "'"
    Else
        strSql = "select FFieldName from ICTemplateEntry where FID='" & strType & "' and FHeadCaption='" & strCaption & "'"
    End If
    
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            CheckFieldName = Trim(rsTemp.Fields("FFieldName").Value)
        Else
            CheckFieldName = ""
        End If
    End If
End Function



