Attribute VB_Name = "MPublic"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'---------------------------------------------------------------------------------------
' Procedure : GetDCSPID
' DateTime  : 2013-1-29 15:20
' Author    :
' Purpose   : 查找仓位信息
'---------------------------------------------------------------------------------------
'
Public Function GetDCSPID(ByVal cnn As ADODB.Connection, ByVal sNumber As String, ByVal lSPGroupID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "SELECT FSPID FROM t_StockPlace  where FSPGroupID =" & lSPGroupID & " and FNumber='" & Trim(sNumber) & "'"
    Set rsTemp = cnn.Execute(strSQL)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetDCSPID = CNulls(rsTemp.Fields("FSPID"), 0)
        End If
    End If
    
    Set rsTemp = Nothing
End Function


''获取仓库组的默认仓位
Public Function GetDEFDCSPID(ByVal cnn As ADODB.Connection, ByVal lSPGroupID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "Select FDefaultSPID from t_StockPlaceGroup where FSPGroupID != 0 And FSPGroupID =" & lSPGroupID
    Set rsTemp = cnn.Execute(strSQL)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            GetDEFDCSPID = CNulls(rsTemp.Fields("FDefaultSPID"), 0)
        End If
    End If
    
    Set rsTemp = Nothing
End Function

''根据条码获取物料ID
Public Function GetICItemID(ByVal cnn As ADODB.Connection, ByVal BarCode As String, ByRef UnitID As Long, ByRef StockID As Long, ByRef Period As Long, ByRef FieldName As String) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select t1.FItemID,FUnitID,FKFPeriod,t2.FItemID GongChang from t_ICItem t1"
    strSQL = strSQL & vbCrLf & "inner join t_Stock t2 on t1." & FieldName & "=t2.FNumber where FBarcode='" & BarCode & "'"
    Set rsTemp = cnn.Execute(strSQL)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            UnitID = CNulls(rsTemp.Fields("FUnitID"), 0)
            StockID = CNulls(rsTemp.Fields("GongChang"), 0)
            Period = CNulls(rsTemp.Fields("FKFPeriod"), 0)
            GetICItemID = CNulls(rsTemp.Fields("FItemID"), 0)
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
Public Function GetIsDCSP(ByVal cnn As ADODB.Connection, ByVal lFItemID As Long, ByRef lSPGroupID As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select FSPGroupID,FIsStockMgr  from t_stock where FItemID=" & lFItemID
    Set rsTemp = cnn.Execute(strSQL)
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
Public Function GetLoactionTypeNum(ByVal cnn As ADODB.Connection, ByVal lStockID As Long) As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select t2.FID,t1.FNumber,t1.FItemID from t_stock t1 " & vbCrLf & _
            "inner join t_submessage t2 on t1.FLocalType=t2.FInterID" & vbCrLf & _
            "Where t1.FItemID = " & lStockID
    Set rsTemp = cnn.Execute(strSQL)
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
Public Function GetStockIDByLoaction(ByVal cnn As ADODB.Connection, ByVal sNumber As String) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select t2.FID,t1.FNumber,t1.FItemID from t_stock t1 " & vbCrLf & _
            "inner join t_submessage t2 on t1.FLocalType=t2.FInterID" & vbCrLf & _
            "Where t2.FID = '" & sNumber & "'"
    Set rsTemp = cnn.Execute(strSQL)
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
Public Function IsExitBill(ByVal cnn As ADODB.Connection, ByVal sBillNo As String, ByVal sTable As String, Optional ByVal lFTranType As Long = 0, Optional ByRef lInterID As Long = 0, Optional sFieldName As String = "FBillNo") As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select FInterID  from " & sTable & " where " & sFieldName & "='" & Trim(sBillNo) & "' and FTranType=" & lFTranType
    Set rsTemp = cnn.Execute(strSQL)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            IsExitBill = True
            lInterID = CNulls(rsTemp.Fields("FInterID"), 0)
        End If
    End If
    Set rsTemp = Nothing
End Function


Public Function GetExitBill(ByVal cnn As ADODB.Connection, ByVal sBillNo As String) As ADODB.Recordset
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select v1.FInterID,v1.FROB,sum(abs(u1.FConsignAmount)) as FSumAmt from ICStockBill v1" & vbCrLf & _
            "inner join ICStockBillEntry u1 on v1.FInterID =u1.FInterID" & vbCrLf & _
            "where v1.FTranType=21 and v1.FPosNum like '" & sBillNo & "%'" & vbCrLf & _
            "group by v1.FROB,v1.FInterID "
    Set rsTemp = cnn.Execute(strSQL)
    
    Set GetExitBill = rsTemp
    Set cnn = Nothing
    Set rsTemp = Nothing
End Function



'---------------------------------------------------------------------------------------
' Procedure : GetItemID
' DateTime  : 2013-1-25 16:50
' Author    :
' Purpose   : 查找基础资料FItemID
'---------------------------------------------------------------------------------------
'
Public Function GetItemID(ByVal cnn As ADODB.Connection, ByVal sNumber As String, ByVal lItemClassID As Long, Optional ByRef lUnitID As Long = 0, Optional ByRef dRate As Double) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    If lItemClassID = 4 Then
        strSQL = "select FUnitID,FItemID,FTaxRate from t_icitem where FNumber='" & Trim(sNumber) & "'"
    Else
        strSQL = "select t2.FName_en,t1.FItemID,t1.fnumber,t1.FName from t_item  t1 " & vbCrLf & _
                "inner join t_itemclass t2 on t1.FItemClassID=t2.FItemClassID" & vbCrLf & _
                "where t1.FDetail=1 and  t1.FshortNumber = '" & sNumber & "' and t1.FItemClassID= " & lItemClassID
    End If
    Set rsTemp = cnn.Execute(strSQL)
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
    Set cnn = Nothing
End Function



'---------------------------------------------------------------------------------------
' Procedure : GetCustomID
' DateTime  : 2013-2-21 00:13
' Author    :
' Purpose   : 根据部门内码取客户
'---------------------------------------------------------------------------------------
'
Public Function GetCustomID(ByVal cnn As ADODB.Connection, ByVal lDeptID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select Flandlordid,* from t_Department where FItemID =" & lDeptID
    Set rsTemp = cnn.Execute(strSQL)
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
Public Function GetSaleType(ByVal cnn As ADODB.Connection, ByVal lItemID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select isnull(FStoreType,0)FStoreType from t_Organization where FItemID=" & lItemID
    Set rsTemp = cnn.Execute(strSQL)
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
    Dim strSQL As String
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
Public Function GetLotInfor(ByVal cnn As ADODB.Connection, ByVal lItemID As Long, ByVal lStockID As Long, ByRef strLotNo As String, ByRef strKFDate As String, ByRef strKFPeriod As String) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo Err
    
    '判断是否有启用批次管理或者保质期管理，如果未启用，则直接返回空值
    strSQL = "select FBatchManager,FISKFPeriod from t_ICItem where FItemID=" & CStr(lItemID)
    Set rsTemp = cnn.Execute(strSQL)
    If Not (rsTemp Is Nothing) Then
        If rsTemp.RecordCount > 0 Then
            If rsTemp.Fields("FBatchManager").Value <> True And rsTemp.Fields("FISKFPeriod").Value <> True Then
                strLotNo = ""
                strKFDate = ""
                strKFPeriod = ""
                GetLotInfor = True
                Set rsTemp = Nothing
                Set cnn = Nothing
                Exit Function
            End If
        End If
    End If
    
    strSQL = "select FItemID,FStockID,FBatchNo,FKFDate,FKFPeriod,FQty from ICInventory"
    strSQL = strSQL & vbCrLf & "Where FItemID=" & CStr(lItemID) & " and FStockID=" & CStr(lStockID)
    strSQL = strSQL & vbCrLf & "group by FItemID,FStockID,FBatchNo,FKFDate,FKFPeriod,FQty order by FItemID,FQty desc"
    Set rsTemp = cnn.Execute(strSQL)
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
    Set cnn = Nothing
    Exit Function
Err:
    GetLotInfor = False
    Set rsTemp = Nothing
    Set cnn = Nothing
    Exit Function
End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckFieldName
' DateTime  : 2014-3-16 16:50
' Author    :
' Purpose   : 查找单据字段名
'---------------------------------------------------------------------------------------
Public Function CheckFieldName(ByVal cnn As ADODB.Connection, strCaption As String, strType As String, iFlag As Integer) As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    'iFlag=0:表头； iFlag=1:表体
    If iFlag = 0 Then
        strSQL = "select FFieldName from ICTemplate where FID='" & strType & "' and FCaption='" & strCaption & "'"
    Else
        strSQL = "select FFieldName from ICTemplateEntry where FID='" & strType & "' and FHeadCaption='" & strCaption & "'"
    End If
    
    Set rsTemp = cnn.Execute(strSQL)
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
    
    lStr = Left(sDsn, InStr(1, sDsn, "Catalog") - 1)
    rStr = Right(sDsn, Len(sDsn) - InStr(1, sDsn, "}") + 1)
    mStr = "Catalog=" & strCatalogName
    strDest = lStr & mStr & rStr
    TransfersDsn = strDest
End Function

'转换导出路径
Public Function TransfersDir(ByVal strDir As String, ByVal strAcSet As String) As String
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim lStr As String
    Dim rStr As String
    Dim mStr As String
    Dim strDest As String
    Dim s() As String
    
    lStr = Left(strDir, InStr(1, strDir, "PROD") + 4)
    rStr = Right(strDir, Len(strDir) - InStr(1, strDir, "\Out\") + 1)
    s() = Split(strAcSet, "_")
    strDest = lStr & s(1) & rStr
    TransfersDir = strDest
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
    lStr = Left(sDsn, InStr(1, sDsn, "Catalog") - 1)
    rStr = Right(sDsn, Len(sDsn) - InStr(1, sDsn, "}") + 1)
    mStr = "Catalog=" & strCatalogName
    strDest = lStr & mStr & rStr
    
    '替换服务器
    lStr = Left(strDest, InStr(1, strDest, "Data Source") - 1)
    rStr = Right(strDest, Len(strDest) - InStr(1, strDest, ";Initial") + 1)
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
'
'Public Sub sqlExt(tmpSql As String)
'Dim cmd As New ADODB.Command
' DbConnect
' Set cmd.ActiveConnection = cnn
' cmd.CommandText = tmpSql
' cmd.Execute
' Set cmd = noting
' Db_Disconnect
'
'End Sub


Public Function ExecSQL(ByVal ssql As String, ByVal dsn As String) As ADODB.Recordset
    Dim obj As Object
    Dim rs As ADODB.Recordset

    Set obj = CreateObject("BillDataAccess.GetData")
    Set rs = obj.ExecuteSQL(dsn, ssql)
    Set obj = Nothing
    Set ExecSQL = rs
    Set rs = Nothing
End Function

Public Sub ExportErrorXML(strFile As String, strError As String)

    Dim xmlDocum As MSXML2.DOMDocument
    Dim xmlRoot As MSXML2.IXMLDOMElement
    Dim xmlNode As MSXML2.IXMLDOMNode
    
    Set xmlDocum = New MSXML2.DOMDocument
    Set xmlRoot = xmlDocum.createElement("ResponseMessage")
    Set xmlDocum.documentElement = xmlRoot
    
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "error", "")
    xmlNode.Text = "true"
    xmlRoot.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "code", "")
    xmlNode.Text = 1
    xmlRoot.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "message", "")
    xmlNode.Text = strError
    xmlRoot.appendChild xmlNode
    
    xmlDocum.Save strFile

End Sub


Public Function ValidateXML(strXML As String, strURL As String, strXSD As String, ByRef strError As String) As Boolean

    Dim xmlSchema As MSXML2.XMLSchemaCache60
    Dim xmlMessage As MSXML2.DOMDocument60
    Dim lngErrCode As Long
    
    Set xmlSchema = New MSXML2.XMLSchemaCache60
    xmlSchema.Add strURL, strXSD
    
    
    Set xmlMessage = New MSXML2.DOMDocument60
    xmlMessage.async = False
    xmlMessage.validateOnParse = True
    xmlMessage.resolveExternals = False
    Set xmlMessage.schemas = xmlSchema
    
    Call xmlMessage.Load(strXML)
    lngErrCode = xmlMessage.Validate()
    If xmlMessage.parseError.errorCode <> 0 Then
        strError = " Reason: " & xmlMessage.parseError.reason
        ValidateXML = False
        Exit Function
    End If
    
    ValidateXML = True

End Function



