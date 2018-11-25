VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSalesDelievery"
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
Public Function CreateSales(ByVal strDSN As String, strUUID As String, ByRef Msg As StringBuilder) As KFO.Dictionary
    Dim strMsg As String
    Dim I As Long
    Dim j As Long
    Dim lngBillInterID As Long
    Dim strInvoiceID As String
    Dim objsave As Object
    Dim objIcInv As Object
    Dim propIcInv As Object
    Dim objtemp As Object
    Dim dctTempHead As Object
    Dim bTrans As Boolean '事务标志
    Dim strDate As String
    Dim strErrorInvoice As String
    Dim rsNewBill  As ADODB.Recordset
    Dim dctReturn As Object
    Dim vctTemp As Object
    Dim dctDetail As Object
    Dim TempSRCInterID As Long
    Dim tempCurrency As String
    Set rsNewBill = New ADODB.Recordset
    Dim StrSql As String
    Dim rs As Recordset
    Dim rsBillsDetail As Recordset
    Dim rsBillsCount As Recordset
    Dim QTY As Long

    
    
    '********需要返回的dictionay
    Set dctReturn = New KFO.Dictionary
    Set vctTemp = New KFO.Vector
    
    '****************************************
    
    On Error GoTo Err_Lab
    Set objsave = CreateObject("BillDataAccess.GetData")
    Set rsNewBill = New ADODB.Recordset
    'Get data 查询一次导入了多少张发票
    StrSql = "select distinct FBillNo from T_t_Sales where FUUID='" & strUUID & "' and ftype = '销售订单'"
    Set rsBillsCount = objsave.ExecuteSQL(strDSN, StrSql)
    
    

    I = 0
    '循环表头添值
    Do While I < rsBillsCount.RecordCount
        StrSql = ""
         StrSql = "select FType,t1.FBillNo SEOBillNo,t2.FProductName,t1.FDate,t1.FWareHouse,t3.FItemID,COUNT(t2.FModel) Count,t2.FModel,t5.FAllAmount,case when FBarSgin=1 then t1.FBarCode else '' end BarCode,"
        StrSql = StrSql & vbCrLf & "case when FBarSgin=0 then t1.FBarCode else '' end BoxCode,t1.fbillno+'-'+FImportTime FImportTime,t4.FCustID,t4.FDeptID,t4.FEmpID,t3.FUnitID,t5.FTaxAmount,t5.FSourceBillNo,t5.FSourceEntryID,t5.FSourceInterId,t1.fuuid from T_t_Sales t1"
        StrSql = StrSql & vbCrLf & "join T_t_Package t2 on t1.FBarCode=t2.FBoxBarCode"
        StrSql = StrSql & vbCrLf & "join t_icitem t3 on t2.FProductNumber = t3.FNumber"
        StrSql = StrSql & vbCrLf & "join SEOrder t4 on t1.FBillNo=t4.FBillNo"
        StrSql = StrSql & vbCrLf & "join SEOrderEntry t5 on t4.FInterID = t5.FInterID"
        StrSql = StrSql & vbCrLf & "where t1.fuuid='" & strUUID & "'"
        StrSql = StrSql & vbCrLf & "group by FType,t1.FBillNo,FProductName,FProductNumber,FBarSgin,t1.FBarCode,FImportTime ,t3.FItemID,t2.FModel,t5.FAllAmount,t1.fuuid,t4.FCustID,t1.FDate,t4.FDeptID,t4.FEmpID,t3.FUnitID,t5.FTaxAmount,t5.FSourceBillNo,t5.FSourceEntryID,t5.FSourceInterId,t1.FWareHouse"
        StrSql = StrSql & vbCrLf & "Union"
        StrSql = StrSql & vbCrLf & "select FType,t1.FBillNo SEOBillNo,t2.FProductName,t1.FDate,t1.FWareHouse,t3.FItemID,COUNT(t2.FModel) Count,'1' as FModel,t5.FAllAmount,case when FBarSgin=1 then t1.FBarCode else '' end BarCode,"
        StrSql = StrSql & vbCrLf & "case when FBarSgin=0 then t1.FBarCode else '' end BoxCode,t1.fbillno+'-'+FImportTime FImportTime,t4.FCustID,t4.FDeptID,t4.FEmpID,t3.FUnitID,t5.FTaxAmount,t5.FSourceBillNo,t5.FSourceEntryID,t5.FSourceInterId,t1.fuuid from T_t_Sales t1"
        StrSql = StrSql & vbCrLf & "join T_t_Package t2 on t1.FBarCode=t2.FHeBarCode"
        StrSql = StrSql & vbCrLf & "join t_icitem t3 on t2.FProductNumber = t3.FNumber"
        StrSql = StrSql & vbCrLf & "join SEOrder t4 on t1.FBillNo=t4.FBillNo"
        StrSql = StrSql & vbCrLf & "join SEOrderEntry t5 on t4.FInterID = t5.FInterID"
        StrSql = StrSql & vbCrLf & "where t1.fuuid='" & strUUID & "'"
        StrSql = StrSql & vbCrLf & "group by FType,t1.FBillNo,FProductName,FProductNumber,FBarSgin,t1.FBarCode,FImportTime ,t3.FItemID,t2.FModel,t5.FAllAmount,t1.fuuid,t4.FCustID,t1.FDate,t4.FDeptID,t4.FEmpID,t3.FUnitID,t5.FTaxAmount,t5.FSourceBillNo,t5.FSourceEntryID,t5.FSourceInterId,t1.FWareHouse"
        Set rsBillsDetail = objsave.ExecuteSQL(strDSN, StrSql)
        
        
        Set rsNewBill = rsBillsDetail
            
        Set objIcInv = CreateObject("BillDataAccess.clsBillPackage")
        '    objIcInv.GetBillEmptyPackage MMTS.PropsString, 86
        objIcInv.GetBillEmptyPackage strDSN, 21
    
        bTrans = False

        strInvoiceID = rsBillsDetail.Fields("SEOBillNo").Value
        
        Set dctTempHead = New KFO.Dictionary
        
        With objIcInv
                'Header
            .SetValue "FBillNO", rsBillsDetail.Fields("SEOBillNo").Value
            
            .SetValue "FDate", Format(rsBillsDetail.Fields("FDate").Value, "yyyy-MM-dd")
            

            .SetValue "FSupplyID", rsBillsDetail.Fields("FCustID").Value


            .SetValue "FDeptID", rsBillsDetail.Fields("FDeptID").Value

            
             .SetValue "FCurrencyID", 1

'             .setValue "FExchangeRate", rsBillsDetail.Fields("FExchangeRate").Value
             
             '往来科目暂时默认为9999
            .SetValue "FCussentAcctid", "9999"
            .SetValue "FFManagerID", "100"
            .SetValue "FSaleStyle", "101"
            .SetValue "FMarketingStyle", "12571"
            .SetValue "FROB", 1
            .SetValue "FBillerID", "16394"
            .SetValue "FCheckerID", "16394"
            

            .SetValue "FSManagerID", rsBillsDetail.Fields("FEmpID").Value
            .SetValue "FFManagerID", rsBillsDetail.Fields("FEmpID").Value

            

            
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
                

                .SetValue "FUnitID", rsNewBill.Fields("FUnitID").Value, False, j
                .SetValue "FItemID", rsNewBill.Fields("FItemID").Value, False, j

                .SetValue "FDCStockID", rsNewBill.Fields("FWareHouse").Value, False, j
                
                QTY = rsNewBill.Fields("Count").Value * rsNewBill.Fields("FModel").Value '数量等于箱数*规格
                
                .SetValue "FAuxQty", QTY, False, j
'                .SetValue "FConsignPrice", rsNewBill.Fields("FUnitPrice").Value, False, j
'                .SetValue "FConsignAmount", rsNewBill.Fields("FTotalAmount").Value, False, j
                .SetValue "FTaxAmount", rsNewBill.Fields("FTaxAmount").Value, False, j
                 .SetValue "FSourceBillNo", rsNewBill.Fields("FSourceBillNo").Value, False, j
                .SetValue "FSourceEntryID", rsNewBill.Fields("FSourceEntryID").Value, False, j
                .SetValue "FSourceInterId", rsNewBill.Fields("FSourceInterId").Value, False, j
                .SetValue "FSourceTranType", 0, False, j
                strInvoiceID = rsNewBill.Fields("SEOBillNo").Value
                rsNewBill.MoveNext
            Next j
             I = I + 1
             rsBillsCount.MoveNext
        End With
        
SaveBill:
        If strMsg <> "" Then
            GoTo Err_Lab
        End If
        
'        GetObjectContext.SetAbort
        bTrans = True '开启事务标志
        
        Set propIcInv = objIcInv.BillPackage
        Set objtemp = objsave.SaveBill(strDSN, propIcInv) 'Save bill
        If objtemp("success") Then                             'Check bill
            lngBillInterID = objtemp("BillInterID")  '返回的单据内码
            
'            If CheckBillBusiLevel(strDSN, 21, lngBillInterID, 16394, 0, 0, objtemp("BillInterID"), strMsg) = False Then
'               strMsg = "Check failed!"
'                GoTo Err_Lab
'            End If
            
        Else
           strMsg = objtemp("ReturnMsg")   '保存失败时返回的出错信息
          
           GoTo Err_Lab
        End If
        
'        GetObjectContext.SetComplete
        bTrans = False '提交事务后，关闭事务标志
        
        GoTo Line_lbl

Line_lbl:
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
Private Function CheckBillBusiLevel(ByVal sdsn As String, _
                ByVal lTransType As Long, _
                ByVal lBillInterid As Long, _
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
    dctPara.Value("PropString") = sdsn
    dctPara.Value("TransType") = lTransType
    dctPara.Value("InterID") = lBillInterid
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




