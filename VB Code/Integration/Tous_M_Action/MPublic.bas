Attribute VB_Name = "MPublic"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : GetExchangeRate
' DateTime  : 2013-1-29 15:20
' Author    :
' Purpose   : 查找汇率
'---------------------------------------------------------------------------------------
Public Function GetExchangeRate(ByVal CNN As ADODB.Connection, ByVal sNumber As String, ByVal lFDate As String, ByRef lFCurrencyID As Long) As Double
Dim strSql As String
Dim rsTemp As ADODB.Recordset
    strSql = "SELECT FCurrencyID FROM t_Currency WHERE FNumber='" & CStr(CNulls(sNumber, "")) & "'"
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            lFCurrencyID = rsTemp("FCurrencyID")
            strSql = " SELECT t3.FExchangeRate FROM t_ExchangeRate t1" & vbCrLf
            strSql = strSql & " INNER JOIN t_ExchangeRate t2 ON t1.FParentID=t2.FID" & vbCrLf
            strSql = strSql & " INNER JOIN t_ExchangeRateEntry t3 ON t1.FID=t3.FID" & vbCrLf
            strSql = strSql & " WHERE t1.FDetail=1 AND t2.FName='公司汇率' " & vbCrLf
            strSql = strSql & " AND t3.FCyTo=" & lFCurrencyID & vbCrLf
            strSql = strSql & " AND t3.FBegDate<='" & lFDate & "' AND t3.FEndDate>='" & lFDate & "'"
            Set rsTemp = CNN.Execute(strSql)
            If Not (rsTemp Is Nothing) Then
                If rsTemp.RecordCount > 0 Then
                    GetExchangeRate = CNulls(rsTemp.Fields("FExchangeRate"), "")
                    Exit Function
                End If
            End If
        End If
    End If
    GetExchangeRate = 0
    Set rsTemp = Nothing
End Function

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

'---------------------------------------------------------------------------------------
' Procedure : GetItemID by Name
' DateTime  : 2013-1-25 16:50
' Author    :
' Purpose   : 查找基础资料FItemID
'---------------------------------------------------------------------------------------
'
Public Function GetItemIDByNumber(ByVal CNN As ADODB.Connection, ByVal sNumber As String, ByVal lItemClassID As Long, Optional ByRef lUnitID As Long = 0, Optional ByRef dRate As Double) As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    If lItemClassID = 4 Then
        strSql = "select FUnitID,FItemID,FTaxRate from t_icitem where FNumber='" & Trim(sNumber) & "'"
    Else
        strSql = "select t2.FName_en,t1.FItemID,t1.fnumber,t1.FName from t_item  t1 " & vbCrLf & _
                "inner join t_itemclass t2 on t1.FItemClassID=t2.FItemClassID" & vbCrLf & _
                "where t1.FDetail=1 and  t1.FNumber = '" & sNumber & "' and t1.FItemClassID= " & lItemClassID
    End If
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetItemIDByNumber = CNulls(rsTemp.Fields("FItemID"), 0)
            If lItemClassID = 4 Then ''如果是物料需查找单位
                lUnitID = CNulls(rsTemp.Fields("FUnitID"), 0)
                dRate = CNulls(rsTemp.Fields("FTaxRate"), 0)
            End If
        End If
    End If
    
    Set rsTemp = Nothing
    Set CNN = Nothing
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

''根据条码获取物料ID
'Public Function GetICItemID(ByVal CNN As ADODB.Connection, ByVal BarCode As String, ByRef UnitID As Long, ByRef FieldName As String) As Long
'    Dim strSQL As String
'    Dim rsTemp As ADODB.Recordset
'
'    strSQL = "select t1.FItemID,FUnitID,FKFPeriod,t2.FItemID GongChang from t_ICItem t1"
'    strSQL = strSQL & vbCrLf & "inner join t_Stock t2 on t1." & FieldName & "=t2.FNumber where FBarcode='" & BarCode & "'"
'    Set rsTemp = CNN.Execute(strSQL)
'    If Not (rsTemp Is Nothing) Then
'        If rsTemp.RecordCount > 0 Then
'            UnitID = CNulls(rsTemp.Fields("FUnitID"), 0)
'            StockID = CNulls(rsTemp.Fields("GongChang"), 0)
'            Period = CNulls(rsTemp.Fields("FKFPeriod"), 0)
'            GetICItemID = CNulls(rsTemp.Fields("FItemID"), 0)
'        End If
'    End If
'
'    Set rsTemp = Nothing
'End Function


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
                "where t1.FDetail=1 and  t1.FNumber = '" & sNumber & "' and t1.FItemClassID= " & lItemClassID
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
                                ByVal lTranstype As Long, _
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
    dctPara.Value("TransType") = lTranstype
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
            
    Set oCheckBill = CreateObject("K3MCheckBill.CheckNow")
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
Public Function GetLotInfor(ByVal CNN As ADODB.Connection, ByVal lItemID As Long, ByVal lStockID As Long, ByRef strLotNo As String, ByRef strKFDate As String, ByRef strKFPeriod As String) As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo Err
    
    '判断是否有启用批次管理或者保质期管理，如果未启用，则直接返回空值
    strSql = "select FBatchManager,FISKFPeriod from t_ICItem where FItemID=" & CStr(lItemID)
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            If rsTemp.Fields("FBatchManager").Value <> True And rsTemp.Fields("FISKFPeriod").Value <> True Then
                strLotNo = ""
                strKFDate = ""
                strKFPeriod = ""
                GetLotInfor = True
                Set rsTemp = Nothing
                Set CNN = Nothing
                Exit Function
            End If
        End If
    End If
    
    strSql = "select FItemID,FStockID,FBatchNo,FKFDate,FKFPeriod,FQty from ICInventory"
    strSql = strSql & vbCrLf & "Where FItemID=" & CStr(lItemID) & " and FStockID=" & CStr(lStockID)
    strSql = strSql & vbCrLf & "group by FItemID,FStockID,FBatchNo,FKFDate,FKFPeriod,FQty order by FItemID,FQty desc"
    Set rsTemp = CNN.Execute(strSql)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 And Val(Trim(rsTemp.Fields("FKFPeriod").Value)) <> 0 Then
            strLotNo = Trim(rsTemp.Fields("FBatchNo").Value)
            strKFDate = Trim(rsTemp.Fields("FKFDate").Value)
            strKFPeriod = Trim(rsTemp.Fields("FKFPeriod").Value)
            GetLotInfor = True
        Else
            strLotNo = "Lot Number"
            strKFDate = "2100-01-01"
            strKFPeriod = "1"
            GetLotInfor = True
        End If
    End If
    
    Set rsTemp = Nothing
    Set CNN = Nothing
    Exit Function
Err:
    GetLotInfor = False
    Set rsTemp = Nothing
    Set CNN = Nothing
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

'转换连接字符串
Public Function TransfersDsn(ByVal strCatalogName As String, ByVal sDsn As String) As String
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim lStr As String
    Dim rStr As String
    Dim mStr As String
    Dim strDest As String
    
    lStr = VBA.Left(sDsn, InStr(1, sDsn, "Catalog") - 1)
    rStr = VBA.Right(sDsn, Len(sDsn) - InStr(1, sDsn, "}") + 1)
    mStr = "Catalog=" & strCatalogName
    strDest = lStr & mStr & rStr
    TransfersDsn = strDest
End Function

'转换连接字符串
Public Function TransfersDsn2(ByVal strDataSource As String, ByVal strCatalogName As String, ByVal strUserName As String, ByVal strPassword As String, ByVal sDsn As String) As String
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim lStr As String
    Dim rStr As String
    Dim mStr As String
    Dim strDest As String
    
    '替换数据库
    lStr = VBA.Left(sDsn, InStr(1, sDsn, "Catalog") - 1)
    rStr = VBA.Right(sDsn, Len(sDsn) - InStr(1, sDsn, "}") + 1)
    mStr = "Catalog=" & strCatalogName
    strDest = lStr & mStr & rStr
    
    '替换服务器
    lStr = VBA.Left(strDest, InStr(1, strDest, "Data Source") - 1)
    rStr = VBA.Right(strDest, Len(strDest) - InStr(1, strDest, ";Initial") + 1)
    mStr = "Data Source=" & strDataSource
    strDest = lStr & mStr & rStr
    
    '替换用户名
    lStr = Left(strDest, InStr(1, strDest, "User ID") - 1)
    rStr = Right(strDest, Len(strDest) - InStr(1, strDest, ";Password") + 1)
    mStr = "User ID=" & strUserName
    strDest = lStr & mStr & rStr
    
    '替换密码
    lStr = Left(strDest, InStr(1, strDest, "Password") - 1)
    rStr = Right(strDest, Len(strDest) - InStr(1, strDest, ";Data") + 1)
    mStr = "Password=" & strPassword
    strDest = lStr & mStr & rStr
    
    
    TransfersDsn2 = strDest
End Function

Public Function ExecSQL1(ByVal ssql As String, ByVal dsn As String) As ADODB.Recordset
    Dim OBJ As Object
    Dim rs As ADODB.Recordset

    Set OBJ = CreateObject("BillDataAccess.GetData")
    Set rs = OBJ.ExecuteSQL(dsn, ssql)
    Set OBJ = Nothing
    Set ExecSQL1 = rs
    Set rs = Nothing
End Function

'根据单据中文名获取老单字段名
Public Function GetKeyField(strFieldName As String, bIsHead As Boolean, strTranType As String, ByRef IsSuccess As Boolean, strDSN As String, DBName As String) As String
Dim strSql As String
Dim rs As ADODB.Recordset
Dim i As Long
On Error GoTo Err
    IsSuccess = False
    
'    strTranType = m_BillTransfer.SaveVect.Item(1).Value("FTransType")
    If bIsHead = True Then
        strSql = "select FFieldName from " & DBName & ".dbo.ICTemplate where FID='" & strTranType & "' and FCaption='" & strFieldName & "'"
    Else
        strSql = "select FFieldName from " & DBName & ".dbo.ICTemplateEntry where FID='" & strTranType & "' and FHeadCaption='" & strFieldName & "'"
    End If
    Set rs = ExecSQL1(strSql, strDSN)
    
    If rs.RecordCount > 0 Then
        GetKeyField = rs.Fields("FFieldName").Value
    Else
        GetKeyField = ""
    End If

    IsSuccess = True
    Set rs = Nothing
    Exit Function
Err:
    Set rs = Nothing
    GetKeyField = ""
End Function

Public Function ExecuteSQL(ByVal ssql As String, ByVal dsn As String) As ADODB.Recordset
    Dim OBJ As Object
    Dim rs As ADODB.Recordset

    Set OBJ = CreateObject("BillDataAccess.GetData")
    Set rs = OBJ.ExecuteSQL(dsn, ssql)
    Set OBJ = Nothing
    Set ExecuteSQL = rs
    Set rs = Nothing
End Function
