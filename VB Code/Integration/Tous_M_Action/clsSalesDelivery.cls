VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSalesDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'********************************************************
'目的:新增销售出库单
'参数：rsNewBill
'返回：无
'Created By Anson 2010-08-23
'********************************************************
Public Function CreateSales(ByVal strDSN As String, strUUID As String, ByRef Msg As StringBuilder, ByVal BillsType As String) As KFO.Dictionary
    Dim strMsg As String
    Dim i As Long
    Dim j As Long
    Dim lngBillInterID As Long
    Dim strInvoiceID As String
    Dim objSave As Object
    Dim objIcInv As Object
    Dim propIcInv As Object
    Dim objtemp As Object
    Dim dctTempHead As Object
    Dim bTrans As Boolean '事务标志
    Dim strDate As String
    Dim strErrorInvoice As String
    Dim rsNewBill  As adodb.Recordset
    Dim dctReturn As Object
    Dim vctTemp As Object
    Dim dctDetail As Object
    Dim TempSRCInterID As Long
    Dim tempCurrency As String
    Set rsNewBill = New adodb.Recordset
    Dim strSql As String
    Dim rs As Recordset
    Dim rsBillsDetail As Recordset
    Dim rsBillsCount As Recordset
    Dim QTY As Long
    
    Set Msg = New StringBuilder

    '********需要返回的dictionay
    Set dctReturn = New KFO.Dictionary
    Set vctTemp = New KFO.Vector
    
    '****************************************
    
    On Error GoTo Err_Lab
    Set objSave = CreateObject("BillDataAccess.GetData")
    Set rsNewBill = New adodb.Recordset
    'Get data 查询一次导入了多少张发票
    strSql = "select distinct FBillNo from T_t_Sales where FUUID='" & strUUID & "' and ftype = '" & BillsType & "'"
    Set rsBillsCount = objSave.ExecuteSQL(strDSN, strSql)
    
    i = 0
    '循环表头添值
    Do While i < rsBillsCount.RecordCount
        
        
            If BillsType = "销售退货" Then
                strSql = "select t1.FType,t1.FBillNo SEOBillNo,CONVERT(varchar(10), t1.FDate, 23)FDate,t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,DATEADD(day,-t3.fkfperiod,t2.FDate) FProductionDate,t3.fkfperiod,"
                strSql = strSql & vbCrLf & "t2.FModel,t4.FInterID FSourceInterId,t5.FEntryID FSourceEntryID,t4.FCustID,t4.FDeptID,t4.FEmpID,t3.FUnitID,"
                strSql = strSql & vbCrLf & "t5.FAuxPrice,sum(t5.FAuxQty) FAuxQty,COUNT(1) FQTY,t8.FInterID FSalesSourceID,t8.FEntryID FSalesSourceEntryID,t7.FBillNo FSalesSourceNo"
                strSql = strSql & vbCrLf & "from T_t_Sales t1 inner join T_t_Package t2 on t1.FBarCode =t2.FBoxBarCode and t1.FBarSgin=1"
                strSql = strSql & vbCrLf & "inner join t_ICItem t3 on t2.FProductNumber=t3.FBarcode"
                strSql = strSql & vbCrLf & "inner join SEOrder t4 on t1.FBillNo=t4.FBillNo"
                strSql = strSql & vbCrLf & "inner join SEOrderEntry t5 on t4.FInterID = t5.FInterID and t5.FItemID=t3.FItemID"
                strSql = strSql & vbCrLf & "inner join t_Stock t6 on t1.FDefaultWareHouse=t6.FItemID"
                strSql = strSql & vbCrLf & "inner join ICStockBillEntry t8 on t4.FInterID=t8.FSourceInterId"
                strSql = strSql & vbCrLf & "inner join ICStockBill t7 on t8.FInterID =t7.FInterID"
                strSql = strSql & vbCrLf & "where t1.fuuid='" & strUUID & "' and t1.FType='" & BillsType & "'  and t6.FNumber not like '%C004%'"
                strSql = strSql & vbCrLf & "group by t1.FType,t1.FBillNo,CONVERT(varchar(10), t1.FDate, 23),t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,t2.FDate,t3.fkfperiod,"
                strSql = strSql & vbCrLf & "t2.FModel , t4.FInterID, t5.FEntryID,t4.FCustID,t4.FDeptID,t4.FEmpID,t3.FUnitID,t5.FAuxPrice,t5.FAuxQty,t8.FInterID,t8.FEntryID,t7.FBillNo "
                strSql = strSql & vbCrLf & "Union"
                strSql = strSql & vbCrLf & "select t1.FType,t1.FBillNo SEOBillNo,CONVERT(varchar(10), t1.FDate, 23)FDate,t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,DATEADD(day,-t3.fkfperiod,t2.FDate) FProductionDate,t3.fkfperiod,"
                strSql = strSql & vbCrLf & "t2.FModel,t4.FInterID FSourceInterId,t5.FEntryID FSourceEntryID,t4.FCustID,t4.FDeptID,t4.FEmpID,t3.FUnitID,"
                strSql = strSql & vbCrLf & "t5.FAuxPrice,sum(t5.FAuxQty) FAuxQty,COUNT(1) FQTY,t8.FInterID FSalesSourceID,t8.FEntryID FSalesSourceEntryID,t7.FBillNo FSalesSourceNo"
                strSql = strSql & vbCrLf & "from T_t_Sales t1 inner join T_t_Package t2 on t1.FBarCode =t2.FHeBarCode and t1.FBarSgin=0"
                strSql = strSql & vbCrLf & "inner join t_ICItem t3 on t2.FProductNumber=t3.FBarcode"
                strSql = strSql & vbCrLf & "inner join SEOrder t4 on t1.FBillNo=t4.FBillNo"
                strSql = strSql & vbCrLf & "inner join SEOrderEntry t5 on t4.FInterID = t5.FInterID and t5.FItemID=t3.FItemID"
                strSql = strSql & vbCrLf & "inner join t_Stock t6 on t1.FDefaultWareHouse=t6.FItemID"
                strSql = strSql & vbCrLf & "inner join ICStockBillEntry t8 on t4.FInterID=t8.FSourceInterId"
                strSql = strSql & vbCrLf & "inner join ICStockBill t7 on t8.FInterID =t7.FInterID"
                strSql = strSql & vbCrLf & "where t1.fuuid='" & strUUID & "' and t1.FType='" & BillsType & "'  and t6.FNumber not like '%C004%'"
                strSql = strSql & vbCrLf & "group by t1.FType,t1.FBillNo,CONVERT(varchar(10), t1.FDate, 23),t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,t2.FDate,t3.fkfperiod,"
                strSql = strSql & vbCrLf & "t2.FModel , t4.FInterID, t5.FEntryID,t4.FCustID,t4.FDeptID,t4.FEmpID,t3.FUnitID,t5.FAuxPrice,t5.FAuxQty,t8.FInterID,t8.FEntryID,t7.FBillNo "
            Else
                strSql = "select t1.FType,t1.FBillNo SEOBillNo,CONVERT(varchar(10), t1.FDate, 23)FDate,t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,DATEADD(day,-t3.fkfperiod,t2.FDate) FProductionDate,t3.fkfperiod,"
                strSql = strSql & vbCrLf & "t2.FModel,t4.FInterID FSourceInterId,t5.FEntryID FSourceEntryID,t4.FCustID,t4.FDeptID,t4.FEmpID,t3.FUnitID,t5.FAuxTaxPrice,t5.FAllAmount,"
                strSql = strSql & vbCrLf & "t5.FAuxPrice,t5.FAuxQty ,t5.FAuxStockQty,COUNT(1) FQTY"
                strSql = strSql & vbCrLf & "from T_t_Sales t1 inner join T_t_Package t2 on t1.FBarCode =t2.FBoxBarCode and t1.FBarSgin=1"
                strSql = strSql & vbCrLf & "inner join t_ICItem t3 on t2.FProductNumber=t3.FBarcode"
                strSql = strSql & vbCrLf & "inner join SEOrder t4 on t1.FBillNo=t4.FBillNo"
                strSql = strSql & vbCrLf & "inner join SEOrderEntry t5 on t4.FInterID = t5.FInterID and t5.FItemID=t3.FItemID"
                strSql = strSql & vbCrLf & "inner join t_Stock t6 on t1.FDefaultWareHouse=t6.FItemID"
                strSql = strSql & vbCrLf & "where t1.fuuid='" & strUUID & "' and t1.FType='" & BillsType & "'  and t6.FNumber not like '%C004%'"
                strSql = strSql & vbCrLf & "group by t1.FType,t1.FBillNo,CONVERT(varchar(10), t1.FDate, 23),t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,t2.FDate,t3.fkfperiod,"
                strSql = strSql & vbCrLf & "t2.FModel , t4.FInterID, t5.FEntryID,t4.FCustID,t4.FDeptID,t4.FEmpID,t3.FUnitID,t5.FAuxPrice,t5.FAuxQty,t5.FAuxTaxPrice,t5.FAllAmount,t5.FAuxStockQty"
                strSql = strSql & vbCrLf & "Union"
                strSql = strSql & vbCrLf & "select t1.FType,t1.FBillNo SEOBillNo,CONVERT(varchar(10), t1.FDate, 23)FDate,t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,DATEADD(day,-t3.fkfperiod,t2.FDate) FProductionDate,t3.fkfperiod,"
                strSql = strSql & vbCrLf & "t2.FModel,t4.FInterID FSourceInterId,t5.FEntryID FSourceEntryID,t4.FCustID,t4.FDeptID,t4.FEmpID,t3.FUnitID,t5.FAuxTaxPrice,t5.FAllAmount,"
                strSql = strSql & vbCrLf & "t5.FAuxPrice,t5.FAuxQty ,t5.FAuxStockQty,COUNT(1) FQTY"
                strSql = strSql & vbCrLf & "from T_t_Sales t1 inner join T_t_Package t2 on t1.FBarCode =t2.FHeBarCode and t1.FBarSgin=0"
                strSql = strSql & vbCrLf & "inner join t_ICItem t3 on t2.FProductNumber=t3.FBarcode"
                strSql = strSql & vbCrLf & "inner join SEOrder t4 on t1.FBillNo=t4.FBillNo"
                strSql = strSql & vbCrLf & "inner join SEOrderEntry t5 on t4.FInterID = t5.FInterID and t5.FItemID=t3.FItemID"
                strSql = strSql & vbCrLf & "inner join t_Stock t6 on t1.FDefaultWareHouse=t6.FItemID"
                strSql = strSql & vbCrLf & "where t1.fuuid='" & strUUID & "' and t1.FType='" & BillsType & "'  and t6.FNumber not like '%C004%'"
                strSql = strSql & vbCrLf & "group by t1.FType,t1.FBillNo,CONVERT(varchar(10), t1.FDate, 23),t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,t2.FDate,t3.fkfperiod,"
                strSql = strSql & vbCrLf & "t2.FModel , t4.FInterID, t5.FEntryID,t4.FCustID,t4.FDeptID,t4.FEmpID,t3.FUnitID,t5.FAuxPrice,t5.FAuxQty,t5.FAuxTaxPrice,t5.FAllAmount,t5.FAuxStockQty"
            
            End If
        
        Set rsBillsDetail = objSave.ExecuteSQL(strDSN, strSql)
        
        
        Set rsNewBill = rsBillsDetail
        
        If rsBillsDetail.RecordCount = 0 Then
            dctReturn("success") = False
            dctReturn("errmsg") = "生成单据时查找数据失败，请检查对应订单是否存在或是订单物料是否一致，或者是否导入相应的包装信息"
            Set CreateSales = dctReturn
            Exit Function
        End If
        
            
        Set objIcInv = CreateObject("BillDataAccess.clsBillPackage")
        '    objIcInv.GetBillEmptyPackage MMTS.PropsString, 86
        objIcInv.GetBillEmptyPackage strDSN, 21
    
        bTrans = False

        strInvoiceID = rsBillsDetail.Fields("SEOBillNo").Value
        
        Set dctTempHead = New KFO.Dictionary
        
        With objIcInv
            'Header
'            .SetValue "FBillNO", rsBillsDetail.Fields("SEOBillNo").Value
            .setValue "FDate", Format(rsBillsDetail.Fields("FDate").Value, "yyyy-MM-dd")
            .setValue "FSupplyID", rsBillsDetail.Fields("FCustID").Value
            .setValue "FDeptID", rsBillsDetail.Fields("FDeptID").Value
'            .SetValue "FCurrencyID", 1

             
             '往来科目暂时默认为9999
'            .SetValue "FCussentAcctid", "9999"
            .setValue "FFManagerID", "100"
            .setValue "FSaleStyle", "101"
            .setValue "FMarketingStyle", "12571"
            
            If BillsType = "销售订单" Then
                .setValue "FROB", 1
            ElseIf BillsType = "销售退货" Then
                .setValue "FROB", -1
            End If
            
                
            .setValue "FBillerID", "16394"
'            .SetValue "FCheckerID", "16394"
            
            .setValue "FSManagerID", rsBillsDetail.Fields("FEmpID").Value
            .setValue "FFManagerID", rsBillsDetail.Fields("FEmpID").Value
        End With
            
            
        With objIcInv                                   'Entry
            For j = 1 To rsNewBill.RecordCount
            
                If rsNewBill.EOF = True Then
                    GoTo SaveBill
                End If
                If rsNewBill.Fields("SEOBillNo").Value <> strInvoiceID Then
                    GoTo SaveBill
                End If
                strInvoiceID = rsNewBill.Fields("SEOBillNo").Value
                .AddRow j
                .setValue "FKFDate", rsNewBill.Fields("FProductionDate").Value, False, j
                .setValue "FKFPeriod", rsNewBill.Fields("fkfperiod").Value, False, j
                .setValue "FBatchNo", rsNewBill.Fields("FBatch").Value, False, j
                .setValue "FUnitID", rsNewBill.Fields("FUnitID").Value, False, j
                .setValue "FItemID", rsNewBill.Fields("FItemID").Value, False, j
                 .setValue "FDCSPID", rsNewBill.Fields("FDefaultSPID").Value, False, j
                .setValue "FDCStockID", rsNewBill.Fields("FDefaultWareHouse").Value, False, j
                
                QTY = rsNewBill.Fields("FQTY").Value * rsNewBill.Fields("FModel").Value '数量等于箱数*规格
                
                .setValue "FAuxQtyMust", rsNewBill.Fields("FQty").Value - rsNewBill.Fields("FAuxStockQty").Value, False, j
                .setValue "FAuxQty", rsNewBill.Fields("FQty").Value, False, j
                .setValue "FConsignPrice", rsNewBill.Fields("FAuxTaxPrice").Value, False, j
                .setValue "FConsignAmount", rsNewBill.Fields("FAllAmount").Value, False, j
'                .SetValue "FTaxAmount", rsNewBill.Fields("FTaxAmount").Value, False, j
                If BillsType = "销售订单" Then
                     .setValue "FSourceBillNo", rsNewBill.Fields("SEOBillNo").Value, False, j
                    .setValue "FSourceEntryID", rsNewBill.Fields("FSourceEntryID").Value, False, j
                    .setValue "FSourceInterId", rsNewBill.Fields("FSourceInterId").Value, False, j
                    .setValue "FSourceTranType", 81, False, j
                Else
                    .setValue "FSourceBillNo", rsNewBill.Fields("FSalesSourceNo").Value, False, j
                    .setValue "FSourceEntryID", rsNewBill.Fields("FSalesSourceEntryID").Value, False, j
                    .setValue "FSourceInterId", rsNewBill.Fields("FSalesSourceID").Value, False, j
                    .setValue "FSourceTranType", 21, False, j
                End If
                    
                    .setValue "FOrderInterID", rsNewBill.Fields("FSourceInterId").Value, j           '
                    .setValue "FOrderEntryID", rsNewBill.Fields("FSourceEntryID").Value, j           '
                    .setValue "FOrderBillNo", rsNewBill.Fields("SEOBillNo").Value, j
                               
                strInvoiceID = rsNewBill.Fields("SEOBillNo").Value
                rsNewBill.MoveNext
            Next j
             
            
        End With
        
SaveBill:
        If strMsg <> "" Then
            GoTo Err_Lab
        End If
        
'        GetObjectContext.SetAbort
        bTrans = True '开启事务标志
        
        Set propIcInv = objIcInv.BillPackage
        Set objtemp = objSave.SaveBill(strDSN, propIcInv) 'Save bill
        If objtemp("success") Then                             'Check bill
            lngBillInterID = objtemp("BillInterID")  '返回的单据内码
            
'            checkBillData strDSN, lngBillInterID, 21, 0, strMsg, 16394 '审核单据
            
            
            '将销售出库单InterID和TranType反写回t_t_sales中
            objSave.ExecuteSQL strDSN, "update t_t_sales set FTransType=21,FIsExported=1,FInterID=" & lngBillInterID & " where FBillNo='" & rsBillsCount.Fields("FBillNo").Value & "' and fuuid='" & strUUID & "'"
            
            If BillsType = "销售订单" Then
                Msg.Append "销售订单号： " & rsBillsCount.Fields("FBillNo").Value & "  生成销售出库单成功，单据编号为： " & objtemp("ReturnBillNo") & vbCrLf
            ElseIf BillsType = "销售退货" Then
                Msg.Append "销售退货单号： " & rsBillsCount.Fields("FBillNo").Value & "  生成销售退货单成功，单据编号为： " & objtemp("ReturnBillNo") & vbCrLf
            End If
            
            rsBillsCount.MoveNext
            i = i + 1
        Else
           strMsg = objtemp("ReturnMsg")   '保存失败时返回的出错信息
           
           If BillsType = "销售订单" Then
            Msg.Append "销售订单号： " & rsBillsCount.Fields("FBillNo").Value & "  生成销售出库单失败" & vbCrLf
           ElseIf BillsType = "销售退货" Then
            Msg.Append "销售退货单号： " & rsBillsCount.Fields("FBillNo").Value & "  生成销售退货单失败" & vbCrLf
           End If
           GoTo Err_Lab
        End If
        
'        GetObjectContext.SetComplete
        bTrans = False '提交事务后，关闭事务标志
        
        GoTo Line_lbl

Line_lbl:
        Set objSave = Nothing
        Set objIcInv = Nothing
    Loop
        dctReturn("success") = True
        dctReturn("errmsg") = strMsg

        Set CreateSales = dctReturn
        
        Exit Function
        
Err_Lab:
        Msg.Append Err.Description
        Msg.Append vbCrLf
        Msg.Append strMsg
        dctReturn("success") = False
        dctReturn("errmsg") = strMsg
        Set CreateSales = dctReturn
        Set objSave = Nothing
End Function

'********************************************************
'目的:审核销售出库单
'参数：sdsn 链接字符串
'      lTransType 单据类型
'       lBillInterid 单据内码
'
'返回：无
'Created By Anson 2010-08-23
'********************************************************
Private Function CheckBillBusiLevel(ByVal sDsn As String, _
                ByVal lTranstype As Long, _
                ByVal lBillInterID As Long, _
                ByVal lUserID As Long, _
                ByVal lCheck As Long, _
                ByVal lBusiLevel As Long, _
                ByVal sBillNo As String, _
                ByRef sRetMsg As String) As Boolean
    Dim dctPara As KFO.Dictionary
    Dim vectCheckItemInfo As KFO.Vector
    Dim oCheckBill As Object
    Dim lRet As Long
    Dim lReturnCode As Long
    
On Error GoTo hrr:
    Set dctPara = New KFO.Dictionary
    dctPara.Value("PropString") = sDsn
    dctPara.Value("TransType") = lTranstype
    dctPara.Value("InterID") = lBillInterID
    dctPara.Value("CheckerID") = lUserID '审核
    dctPara.Value("CheckSwitch") = lCheck
    If lCheck = 0 Then '
        dctPara.Value("OperateCode") = 1
    Else
        dctPara.Value("OperateCode") = 2
    End If
    dctPara.Value("ReCheck") = 1
    dctPara.Value("Operatetype") = 0
    dctPara.Value("CheckDate") = Date
    dctPara.Value("ReturnCode") = 2
    dctPara.Value("ReturnString") = ""
    Set vectCheckItemInfo = New KFO.Vector
    Set dctPara.Value("vectItemInfo") = vectCheckItemInfo
    Set vectCheckItemInfo = Nothing
    dctPara.Value("MultiCheckLevel") = lBusiLevel
    dctPara.Value("WorkFlowFlag") = 0
    dctPara.Value("BillNo") = sBillNo
            
    Set oCheckBill = CreateObject("K3MCheckBill.CheckNow")
    lRet = oCheckBill.CheckBill(dctPara)
    Set oCheckBill = Nothing
    
    lReturnCode = dctPara.GetValue("ReturnCode", 0)
    
    If lReturnCode = 0 Then
        CheckBillBusiLevel = True
    Else
        CheckBillBusiLevel = False
        sRetMsg = "(RetCode:" & lReturnCode & ")"
        sRetMsg = dctPara.GetValue("ReturnString", 0)
        GoTo hrr
    End If
    Set dctPara = Nothing
    Set vectCheckItemInfo = Nothing
    Set oCheckBill = Nothing
    Exit Function
hrr:
    If sRetMsg = "" Then sRetMsg = Err.Description
    Set dctPara = Nothing
    Set vectCheckItemInfo = Nothing
    Set oCheckBill = Nothing
    CheckBillBusiLevel = False
End Function






