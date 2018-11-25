VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "Action_SyncSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private datasource As CDataSource
Private m_oParse As CParse
Private strLanguage As String

 
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

'---------------------------------------------------------------------------------------
' Procedure : SynPackage
' DateTime  : 2018-9-5
' Author    :
' Purpose   : 生成销售出库单并审核
'---------------------------------------------------------------------------------------
'
Public Function SyncSalesDelivery(ByVal sDsn As String, strTablename As String, ByRef strInterIDs As String, ByRef sBillNOs As String, ByRef sErr As String) As Boolean
 
    Dim CNN As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim clsBill As AutoOldBill
    Dim bRet As Boolean
    Dim lRow As Long
    Dim i, j As Long
    
    Dim lFItemID As Long
    Dim lUnitID As Long
    Dim lROB As Long
    
    Dim strBillNo As String
    Dim lInterID As Long
    Dim strSQL As String
    Dim Row As Long
    Dim rs As ADODB.Recordset
    Dim rsBill As ADODB.Recordset
    Dim rsBillInfo As ADODB.Recordset
    Dim bAutoInv As Boolean
    
On Error GoTo SynBillB01_Error
    
    '测试用
'    sDsn = "ConnectString={Persist Security Info=True;Provider=SQLOLEDB.1;User ID=sa;Password=sa;Data Source=.;Initial Catalog=Bluesea};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=BAD160C5-B4DB-4965-A8A7-0D373C885AE6"
'    strTablename = "t_POS_SAL"
    
    InitDataEnv sDsn
    Set CNN = datasource.Connection
    
    '获取本次导入文件中的单据信息，如果为空则报错
    Set rsBill = ExecSQL("select distinct FPOSNumber,FSalesType from  " & strTablename & " order by FPOSNumber", sDsn)
    If rsBill.RecordCount <= 0 Then
        sErr = "There is no data in the EDI file, please contact the administrator for double checking."
        GoTo SynBillB01_Error
    End If
    
    '获取导入明细，准备进行导入
    strSQL = "select t1.*,isnull(t2.FItemID,0) FCustID,t2.FAutoInv, isnull(t3.FItemID,0) FDeptID, isnull(t4.FItemID,0) FItemID,isnull(t4.FUnitID,0) FUnitID, "
    strSQL = strSQL & vbCrLf & "isnull(t5.FItemID,0) FStockID,isnull(t6.FItemID,0) FEmpID from " & strTablename & " t1 "
    strSQL = strSQL & vbCrLf & "left join t_Organization t2 on t1.FStoreCode = t2.FNumber"
    strSQL = strSQL & vbCrLf & "left join t_Department t3 on t1.FStoreCode=t3.FNumber"
    strSQL = strSQL & vbCrLf & "left join t_ICItem t4 on t1.FSKU = t4.FNumber"
    strSQL = strSQL & vbCrLf & "left join t_Stock t5 on t1.FStoreCode = t5.FNumber"
    strSQL = strSQL & vbCrLf & "left join t_Emp t6 on t1.FStoreCode = t6.FNumber"
    strSQL = strSQL & vbCrLf & "order by t1.FStoreCode,t1.FPOSNumber,t1.FLineNumber"
    
    Set rsBillInfo = ExecSQL(strSQL, sDsn)
            
    Set clsBill = New AutoOldBill
    
    lRow = 1
    
    With clsBill
        .sDsn = sDsn
                
        For i = 1 To rsBill.RecordCount
        
            bRet = .NewBill(21)
    
            If bRet = True Then
            
                rsBillInfo.Filter = "FPOSNumber='" & rsBill.Fields("FPOSNumber") & "' and FSalesType='" & rsBill.Fields("FSalesType") & "'"
                bAutoInv = CNulls(rsBillInfo.Fields("FAutoInv").Value, 0)
            
                '检查POS Number是否存在，存在则不允许重复导入
                Set rs = ExecSQL("select 1 from ICStockBill where FTranType=21 and isnull(FPOSNumber,'')='" & rsBill.Fields("FPOSNumber") & "' and FROB=" & rsBill.Fields("FSalesType"), sDsn)
                If rs.RecordCount > 0 Then
                    sErr = "POS Number[" & rsBill.Fields("FPOSNumber").Value & "] has existed in ERP system."
                    GoTo SynBillB01_Error
                End If
                            
                .SetHeadValue "FDate", CDate(rsBillInfo.Fields("FDate").Value)  '日期
                .SetHeadValue "FPOSNumber", CStr(rsBillInfo.Fields("FPOSNumber").Value)  'POS Number
                
                .SetHeadValue "FSaleStyle", 101 '销售方式:默认赊销
                .SetHeadValue "FMarketingStyle", 12530 '销售业务类型:默认销售出库
                
                .SetHeadValue "FROB", Val(rsBill.Fields("FSalesType")) '红蓝字标识
                lROB = Val(rsBill.Fields("FSalesType"))
                
                '检查客户是否存在
                If Val(rsBillInfo.Fields("FCustID").Value) = 0 Then
                    sErr = "Customer Master Data [" & rsBillInfo.Fields("FStoreCode").Value & "] is NOT created in ERP system."
                    GoTo SynBillB01_Error
                End If
                .SetHeadValue "FSupplyID", rsBillInfo.Fields("FCustID").Value
            
                '检查部门是否存在
                If Val(rsBillInfo.Fields("FDeptID").Value) = 0 Then
                    sErr = "Department Master Data [" & rsBillInfo.Fields("FStoreCode").Value & "] is NOT created in ERP system."
                    GoTo SynBillB01_Error
                End If
                .SetHeadValue "FDeptID", rsBillInfo.Fields("FDeptID")
                
                '检查职员是否存在
                If Val(rsBillInfo.Fields("FEmpID").Value) = 0 Then
                    sErr = "Employee Master Data [" & rsBillInfo.Fields("FStoreCode").Value & "] is NOT created in ERP system."
                    GoTo SynBillB01_Error
                End If
                .SetHeadValue "FFManagerID", rsBillInfo.Fields("FEmpID")
                .SetHeadValue "FSManagerID", rsBillInfo.Fields("FEmpID")

            
                '表体赋值
                For j = 1 To rsBillInfo.RecordCount
                                
                    .AddRow lRow
                    
                    
                    Row = lRow
    
                     '检查物料是否存在
                     If Val(rsBillInfo.Fields("FItemID").Value) = 0 Then
                         sErr = "Material Master Data [" & rsBillInfo.Fields("FSKU").Value & "] is NOT created in ERP system."
                         GoTo SynBillB01_Error
                     End If
                    .SetEntryValue "FItemID", rsBillInfo.Fields("FItemID").Value, lRow            '物料
                    .SetEntryValue "FUnitID", rsBillInfo.Fields("FUnitID").Value, lRow            '单位
                    
                    '数量必须大于0
                    If Val(rsBillInfo.Fields("FQty").Value) <= 0 Then
                        sErr = "The quantity can NOT be zero."
                        GoTo SynBillB01_Error
                    End If
                    .SetEntryValue "Fauxqty", Val(rsBillInfo.Fields("FQty").Value) * lROB, lRow           '数量
                    .SetEntryValue "FAuxQtyMust", Val(rsBillInfo.Fields("FQty").Value) * lROB, lRow          '应发数量
                    
                     If Val(rsBillInfo.Fields("FStockID").Value) = 0 Then
                         sErr = "Warehouse Master Data [" & rsBillInfo.Fields("FStoreCode").Value & "] is NOT created in ERP system."
                         GoTo SynBillB01_Error
                     End If
                    .SetEntryValue "FDCStockID", rsBillInfo.Fields("FStockID").Value, lRow        '发货仓库
                
                    .SetEntryValue "FConsignPrice", CNulls(rsBillInfo.Fields("FPrice").Value, 0), lRow           ''单价
                    .SetEntryValue "FDiscountAmount", CNulls(rsBillInfo.Fields("FDiscountAmount").Value, 0) * lROB, lRow          ''折扣额
                    .SetEntryValue "FDiscountRate", CNulls(rsBillInfo.Fields("FDiscountRate").Value, 0), lRow           ''折扣率
                    .SetEntryValue "FConsignAmount", CNulls(rsBillInfo.Fields("FTotalAmount").Value, 0) * lROB, lRow          ''销售金额
                    
                    
                    .SetEntryValue "FNote", rsBillInfo.Fields("FNote").Value, lRow          ''备注
                

                '=========源单信息===========
'                .SetEntryValue "FSourceTranType", "81", lRow          ''源单类型
'                .SetEntryValue "FSourceBillNo", rsBillInfo.Fields("FBillNo").Value, lRow          ''源单编号
'                .SetEntryValue "FSourceInterId", rsBillInfo.Fields("FInterID").Value, lRow          ''源单内码
'                .SetEntryValue "FSourceEntryID", rsBillInfo.Fields("FEntryID").Value, lRow          ''源单行号
'                .SetEntryValue "FOrderInterID", rsBillInfo.Fields("FInterID").Value, lRow            '
'                .SetEntryValue "FOrderEntryID", rsBillInfo.Fields("FEntryID").Value, lRow           '
'                .SetEntryValue "FOrderBillNo", rsBillInfo.Fields("FBillNo").Value, lRow
'                .SetEntryValue "FContractBillNo", rsBillInfo.Fields("FSourceBillNo").Value, lRow
'                .SetEntryValue "FContractEntryID", rsBillInfo.Fields("FSourceEntryID").Value, lRow
'                .SetEntryValue "FContractInterID", rsBillInfo.Fields("FSourceInterId").Value, lRow
               '=============================
               

                    lRow = lRow + 1
NextRow:
                    rsBillInfo.MoveNext
                Next j
NewRow:
                lInterID = 0
                strBillNo = ""
                bRet = .SaveBill(lInterID, strBillNo, sErr)
                If bRet = True Then
                    If checkBillData(sDsn, lInterID, 21, 0, sErr, 16394) = True And bAutoInv = True Then '审核单据
                        
                        If GenerateSI(sDsn, lInterID, 16394, sErr) = False Then
                            GoTo SynBillB01_Error
                        End If
                    
                        strInterIDs = strInterIDs & CStr(lInterID) & ","
                        sBillNOs = sBillNOs & strBillNo & ","
                    Else
                        sErr = "Approve the sales delivery failed: " & sErr
                        GoTo SynBillB01_Error
                    End If
                End If
                
                
                rsBill.MoveNext
                rsBillInfo.Filter = ""
                bAutoInv = False
            
            End If
        Next i
            
            
    End With
    
    Set clsBill = Nothing
    Set CNN = Nothing
    SyncSalesDelivery = True
    GetObjectContext.SetComplete

    Exit Function

SynBillB01_Error:
    SyncSalesDelivery = False
'    GetObjectContext.SetAbort
    Set clsBill = Nothing
    Set CNN = Nothing
    sErr = "Generate Sales Delivery failed：" & sErr & Err.Description
    
End Function

Private Function ExecSQL(ByVal ssql As String, ByVal dsn As String) As ADODB.Recordset
    Dim OBJ As Object
    Dim rs As ADODB.Recordset

    Set OBJ = CreateObject("BillDataAccess.GetData")
    Set rs = OBJ.ExecuteSQL(dsn, ssql)
    Set OBJ = Nothing
    Set ExecSQL = rs
    Set rs = Nothing
End Function


Public Function Test(ByVal s As String) As Boolean
    If s = "a" Then
        Test = True
    Else
        Test = False
    End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : SynSalesInvoice
' DateTime  : 2018-9-5
' Author    :
' Purpose   : 如果客户基础资料勾选了自动生成发票，生成销售发票并审核
'---------------------------------------------------------------------------------------
'
Private Function GenerateSI(ByVal sDsn As String, ByVal lInterID As Long, ByVal lUserID As Long, ByRef sErr As String) As Boolean
 
    Dim CNN As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim clsBill As AutoOldBill
    Dim bRet As Boolean
    Dim lRow As Long
    Dim i, j As Long
    Dim lFItemID As Long
    Dim lUnitID As Long
    Dim strSQL As String
    Dim strBillNo As String
    
    Dim lROB As Long, Row As Long
    
On Error GoTo SynBillB01_Error
        
    InitDataEnv sDsn
    Set CNN = datasource.Connection
    
    '获取出库单明细，准备生成发票
    strSQL = "select t1.FInterID,t1.FROB,t1.FBillNo,t1.FPOSNumber,t1.FSupplyID FCustID,t1.FDate,t1.FSaleStyle,t1.FDeptID,t1.FFManagerID FEmpID,"
    strSQL = strSQL & vbCrLf & "t2.FEntryID, t2.fItemID , t2.FUnitID, t2.FConsignPrice, t2.FQty, t2.FDiscountRate, t2.FDiscountAmount, t2.FConsignAmount,"
    strSQL = strSQL & vbCrLf & "t2.FNote, t3.FTaxRate, t3.FPriceDecimal from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
    strSQL = strSQL & vbCrLf & "inner join t_ICItem t3 on t2.FItemID=t3.FItemID where t1.FInterID=" & CStr(lInterID)
    '无单价商品（包括赠品）不开票
    strSQL = strSQL & vbCrLf & "and t2.FConsignAmount<>0"
    strSQL = strSQL & vbCrLf & "order by t1.FInterID,t2.FEntryID"
    
    Set rs = ExecSQL(strSQL, sDsn)
            
    Set clsBill = New AutoOldBill
    
    lRow = 1
    
    With clsBill
        .sDsn = sDsn
        
        bRet = .NewBill(80)

        If bRet = True Then
                        
            .SetHeadValue "FCustID", rs.Fields("FCustID").Value '客户
            .SetHeadValue "FDeptID", rs.Fields("FDeptID") '部门
            .SetHeadValue "FEmpID", rs.Fields("FEmpID") '业务员
            .SetHeadValue "FDate", CDate(rs.Fields("FDate").Value)  '日期
            .SetHeadValue "FPOSNumber", CStr(rs.Fields("FPOSNumber").Value)  'POS Number
            .SetHeadValue "FSaleStyle", rs.Fields("FSaleStyle").Value '销售方式
'            .SetHeadValue "FMarketingStyle", rs.Fields("FSaleStyle").Value '销售业务类型:默认销售出库
            
            .SetHeadValue "FROB", Val(rs.Fields("FROB")) '红蓝字标识
            lROB = Val(rs.Fields("FROB"))
            
            .SetHeadValue "FCurrencyID", 1   '币别：默认本位币
            .SetHeadValue "FExchangeRateType", 1    '汇率类型：默认公司汇率
            .SetHeadValue "FExchangeRate", 1    '汇率：默认1

        
            '表体赋值
            For j = 1 To rs.RecordCount
                
                            
                .AddRow lRow
                
                Row = lRow

                .SetEntryValue "FItemID", rs.Fields("FItemID").Value, lRow            '物料
                .SetEntryValue "FUnitID", rs.Fields("FUnitID").Value, lRow            '单位
                
                .SetEntryValue "Fauxqty", Val(rs.Fields("FQty").Value) * lROB, lRow           '数量
                .SetEntryValue "Fqty", Val(rs.Fields("FQty").Value) * lROB, lRow          '基本单位数量
            
                .SetEntryValue "FAuxTaxPrice", Val(rs.Fields("FConsignPrice").Value), lRow           '含税单价
                .SetEntryValue "FAuxPriceDiscount", Val(rs.Fields("FConsignAmount").Value) / Val(rs.Fields("FQty").Value), lRow         '实际含税单价
                .SetEntryValue "Fauxprice", KDRound((Val(rs.Fields("FConsignPrice").Value) / (1 + Val(rs.Fields("FTaxRate").Value) / 100)), Val(rs.Fields("FPriceDecimal").Value)), lRow     '含税单价
                
                .SetEntryValue "FAmtDiscount", Val(rs.Fields("FDiscountAmount").Value) * lROB, lRow          ''折扣额
                .SetEntryValue "FStdAmtDiscount", Val(rs.Fields("FDiscountAmount").Value) * lROB, lRow          ''折扣额（本位币）
                .SetEntryValue "FDiscountRate", Val(rs.Fields("FDiscountRate").Value), lRow           ''折扣率
                
                .SetEntryValue "FTaxRate", Val(rs.Fields("FTaxRate").Value), lRow         '税率
                .SetEntryValue "FTaxAmount", KDRound(Val(rs.Fields("FConsignAmount").Value) - (Val(rs.Fields("FConsignAmount").Value) / (1 + Val(rs.Fields("FTaxRate").Value) / 100)), 2) * lROB, lRow        '税额
                .SetEntryValue "FStdTaxAmount", KDRound(Val(rs.Fields("FConsignAmount").Value) - (Val(rs.Fields("FConsignAmount").Value) / (1 + Val(rs.Fields("FTaxRate").Value) / 100)), 2) * lROB, lRow       '税额（本位币）
                
                .SetEntryValue "FAllAmount", Val(rs.Fields("FConsignAmount").Value) * lROB, lRow          '价税合计
                .SetEntryValue "FStdAllAmount", Val(rs.Fields("FConsignAmount").Value) * lROB, lRow       '价税合计（本位币）
                .SetEntryValue "FAmount", KDRound((Val(rs.Fields("FConsignAmount").Value) / (1 + Val(rs.Fields("FTaxRate").Value) / 100)), 2) * lROB, lRow         '金额
                .SetEntryValue "FStdAmount", KDRound((Val(rs.Fields("FConsignAmount").Value) / (1 + Val(rs.Fields("FTaxRate").Value) / 100)), 2) * lROB, lRow        '金额（本位币）
                
                .SetEntryValue "FNote", rs.Fields("FNote").Value, lRow          ''备注
            

            '=========源单信息===========
                .SetEntryValue "FSourceTranType", "21", lRow          ''源单类型
                .SetEntryValue "FSourceBillNo", rs.Fields("FBillNo").Value, lRow          ''源单编号
                .SetEntryValue "FSourceInterId", rs.Fields("FInterID").Value, lRow          ''源单内码
                .SetEntryValue "FSourceEntryID", rs.Fields("FEntryID").Value, lRow          ''源单行号
'                .SetEntryValue "FOrderInterID", rs.Fields("FInterID").Value, lRow            '
'                .SetEntryValue "FOrderEntryID", rs.Fields("FEntryID").Value, lRow           '
'                .SetEntryValue "FOrderBillNo", rs.Fields("FBillNo").Value, lRow
'                .SetEntryValue "FContractBillNo", rs.Fields("FSourceBillNo").Value, lRow
'                .SetEntryValue "FContractEntryID", rs.Fields("FSourceEntryID").Value, lRow
'                .SetEntryValue "FContractInterID", rs.Fields("FSourceInterId").Value, lRow
           '=============================
           

                lRow = lRow + 1
NextRow:
                rs.MoveNext
            Next j
NewRow:
            lInterID = 0
            strBillNo = ""
            bRet = .SaveBill(lInterID, strBillNo, sErr)
            If bRet = True Then
                checkBillData sDsn, lInterID, 80, 0, sErr, 16394 '审核单据
            End If
            
        
        End If
            
            
    End With
    
    Set clsBill = Nothing
    Set CNN = Nothing
    GenerateSI = True

    Exit Function

SynBillB01_Error:
    GenerateSI = False
    Set clsBill = Nothing
    Set CNN = Nothing
    sErr = "Generate Sales Invoice error：" & sErr & Err.Description
    
End Function


