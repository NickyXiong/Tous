Attribute VB_Name = "Main"
Public Function ShowFrm()
Dim frm As Form
Set frm = New frmExport
If frm.setfunc Then frm.Show 0
Set frm = Nothing
End Function



'---------------------------------------------------------------------------------------
' Procedure : GetDCSPID
' DateTime  : 2013-1-29 15:20
' Author    :
' Purpose   : ���Ҳ�λ��Ϣ
'---------------------------------------------------------------------------------------
'
Public Function GetDCSPID(ByVal CNN As adodb.Connection, ByVal sNumber As String, ByVal lSPGroupID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
    strSQL = "SELECT FSPID FROM t_StockPlace  where FSPGroupID =" & lSPGroupID & " and FNumber='" & Trim(sNumber) & "'"
    Set rsTemp = CNN.Execute(strSQL)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetDCSPID = CNulls(rsTemp.Fields("FSPID"), 0)
        End If
    End If
    
    Set rsTemp = Nothing
End Function


''��ȡ�ֿ����Ĭ�ϲ�λ
Public Function GetDEFDCSPID(ByVal CNN As adodb.Connection, ByVal lSPGroupID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
    strSQL = "Select FDefaultSPID from t_StockPlaceGroup where FSPGroupID != 0 And FSPGroupID =" & lSPGroupID
    Set rsTemp = CNN.Execute(strSQL)
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
' Purpose   : ����Ƿ���в�λ����
'---------------------------------------------------------------------------------------
'
Public Function GetIsDCSP(ByVal CNN As adodb.Connection, ByVal lFItemID As Long, ByRef lSPGroupID As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
    strSQL = "select FSPGroupID,FIsStockMgr  from t_stock where FItemID=" & lFItemID
    Set rsTemp = CNN.Execute(strSQL)
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
' Purpose   : ���Ҳֿ�����locationtype number �ֶ�ֵ
'---------------------------------------------------------------------------------------
'
Public Function GetLoactionTypeNum(ByVal CNN As adodb.Connection, ByVal lStockID As Long) As String
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
    strSQL = "select t2.FID,t1.FNumber,t1.FItemID from t_stock t1 " & vbCrLf & _
            "inner join t_submessage t2 on t1.FLocalType=t2.FInterID" & vbCrLf & _
            "Where t1.FItemID = " & lStockID
    Set rsTemp = CNN.Execute(strSQL)
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
' Purpose   : ����locationtypeֵȡ�м�ֵĲֿ�
'---------------------------------------------------------------------------------------
'
Public Function GetStockIDByLoaction(ByVal CNN As adodb.Connection, ByVal sNumber As String) As Long
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
    strSQL = "select t2.FID,t1.FNumber,t1.FItemID from t_stock t1 " & vbCrLf & _
            "inner join t_submessage t2 on t1.FLocalType=t2.FInterID" & vbCrLf & _
            "Where t2.FID = '" & sNumber & "'"
    Set rsTemp = CNN.Execute(strSQL)
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
' Purpose   : ��ѯ�����Ƿ����
'---------------------------------------------------------------------------------------
'
Public Function IsExitBill(ByVal CNN As adodb.Connection, ByVal sBillNo As String, ByVal sTable As String, Optional ByVal lFTranType As Long = 0, Optional ByRef lInterID As Long = 0, Optional sFieldName As String = "FBillNo") As Boolean
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
    strSQL = "select FInterID  from " & sTable & " where " & sFieldName & "='" & Trim(sBillNo) & "' and FTranType=" & lFTranType
    Set rsTemp = CNN.Execute(strSQL)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            IsExitBill = True
            lInterID = CNulls(rsTemp.Fields("FInterID"), 0)
        End If
    End If
    Set rsTemp = Nothing
End Function


Public Function GetExitBill(ByVal CNN As adodb.Connection, ByVal sBillNo As String) As adodb.Recordset
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
    strSQL = "select v1.FInterID,v1.FROB,sum(abs(u1.FConsignAmount)) as FSumAmt from ICStockBill v1" & vbCrLf & _
            "inner join ICStockBillEntry u1 on v1.FInterID =u1.FInterID" & vbCrLf & _
            "where v1.FTranType=21 and v1.FPosNum like '" & sBillNo & "%'" & vbCrLf & _
            "group by v1.FROB,v1.FInterID "
    Set rsTemp = CNN.Execute(strSQL)
    
    Set GetExitBill = rsTemp
    Set CNN = Nothing
    Set rsTemp = Nothing
End Function



'---------------------------------------------------------------------------------------
' Procedure : GetItemID
' DateTime  : 2013-1-25 16:50
' Author    :
' Purpose   : ���һ�������FItemID
'---------------------------------------------------------------------------------------
'
Public Function GetItemID(ByVal CNN As adodb.Connection, ByVal sNumber As String, ByVal lItemClassID As Long, Optional ByRef lUnitID As Long = 0, Optional ByRef dRate As Double) As Long
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    If lItemClassID = 4 Then
        strSQL = "select FUnitID,FItemID,FTaxRate from t_icitem where FNumber='" & Trim(sNumber) & "'"
    Else
        strSQL = "select t2.FName_en,t1.FItemID,t1.fnumber,t1.FName from t_item  t1 " & vbCrLf & _
                "inner join t_itemclass t2 on t1.FItemClassID=t2.FItemClassID" & vbCrLf & _
                "where t1.FDetail=1 and  t1.FshortNumber like '%" & sNumber & "' and t1.FItemClassID= " & lItemClassID
    End If
    Set rsTemp = CNN.Execute(strSQL)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetItemID = CNulls(rsTemp.Fields("FItemID"), 0)
            If lItemClassID = 4 Then ''�������������ҵ�λ
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
' Purpose   : ���ݲ�������ȡ�ͻ�
'---------------------------------------------------------------------------------------
'
Public Function GetCustomID(ByVal CNN As adodb.Connection, ByVal lDeptID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
    strSQL = "select Flandlordid,* from t_Department where FItemID =" & lDeptID
    Set rsTemp = CNN.Execute(strSQL)
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
' Purpose   : ���ҿͻ�����Store Typeֵ
'---------------------------------------------------------------------------------------
'
Public Function GetSaleType(ByVal CNN As adodb.Connection, ByVal lItemID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
    strSQL = "select isnull(FStoreType,0)FStoreType from t_Organization where FItemID=" & lItemID
    Set rsTemp = CNN.Execute(strSQL)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetSaleType = CNulls(rsTemp.Fields("FStoreType"), 0)
        End If
    End If
    Set rsTemp = Nothing
End Function




'/������˴���lCheck=0����ʾ��ˣ�lCheck=1����ʾ����ˣ�lStockType�ֿ����� 0����ʵ�� ��1�������
Public Function checkBillData(ByVal sDsn As String, _
                                ByVal lBillInterID As Long, _
                                ByVal lTransType As Long, _
                                ByVal lCheck As Long, _
                                ByRef sRetMsg As String, _
                                Optional ByVal lcheckID As Long = 0, _
                                Optional ByVal lStockType As Long = 0) As Boolean
'On Error GoTo HError
    Dim oCheckBill As Object
    Dim strSQL As String
    Dim lRet As Long
    
    Dim rs As adodb.Recordset
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
    '�����Ƿ�����ʵ�ָ�����ж�
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
' Purpose   : ���һ�������FItemID
'---------------------------------------------------------------------------------------
'
Public Function GetLotInfor(ByVal lItemID As Long, ByVal lStockID As Long, ByVal lSPID As Long, ByRef strLotNo As String, ByRef strKFDate As String, ByRef strKFPeriod As String) As Boolean
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
On Error GoTo Err
    
    '�ж��Ƿ����������ι�����߱����ڹ������δ���ã���ֱ�ӷ��ؿ�ֵ
    strSQL = "select FBatchManager,FISKFPeriod from t_ICItem where FItemID=" & CStr(lItemID)
    Set rsTemp = modPub.ExecSQL(strSQL)
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
    
    strSQL = "select top 1 FBatchNo,FKFDate,FKFPeriod,FQty from ICInventory  "
    strSQL = strSQL & vbCrLf & "Where FQty<>0 and FItemID=" & CStr(lItemID) & " and FStockID=" & CStr(lStockID) & " and FStockPlaceID=" & CStr(lSPID)
    strSQL = strSQL & vbCrLf & "group by FBatchNo,FKFDate,FKFPeriod,FQty"
    strSQL = strSQL & vbCrLf & "order by FKFDate"
    Set rsTemp = modPub.ExecSQL(strSQL)
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
' Purpose   : ���ҵ����ֶ���
'---------------------------------------------------------------------------------------
Public Function CheckFieldName(ByVal CNN As adodb.Connection, strCaption As String, strType As String, iFlag As Integer) As String
    Dim strSQL As String
    Dim rsTemp As adodb.Recordset
    
    'iFlag=0:��ͷ�� iFlag=1:����
    If iFlag = 0 Then
        strSQL = "select FFieldName from ICTemplate where FID='" & strType & "' and FCaption='" & strCaption & "'"
    Else
        strSQL = "select FFieldName from ICTemplateEntry where FID='" & strType & "' and FHeadCaption='" & strCaption & "'"
    End If
    
    Set rsTemp = CNN.Execute(strSQL)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            CheckFieldName = Trim(rsTemp.Fields("FFieldName").Value)
        Else
            CheckFieldName = ""
        End If
    End If
End Function



