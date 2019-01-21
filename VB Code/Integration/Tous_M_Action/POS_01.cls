VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************
'Function: Import Sales Transaction
'Date: 2013/11/08
'*******************************************************

Option Explicit

Private datasource As CDataSource
Private m_oParse As CParse
Private strLanguage As String


'/* 销售出库单 Sales
Private Const CONST_B01_FBillNo As Long = 0 '销售店铺

'Added by Nicky  - 20150504
'增加是否销售类型字段
Private Const CONST_B01_FType As Long = 1 '销售出库单代码
'Added by Nicky  - 20150504

Private Const CONST_B01_FDATE As Long = 2 '销售出库单代码
Private Const CONST_B01_FEntryID  As Long = 3 '销售/退货
Private Const CONST_B01_FItemNumber  As Long = 4 '日期
Private Const CONST_B01_FUOM As Long = 5 '物料代码
Private Const CONST_B01_FLotNumber As Long = 6 ' 数量
Private Const CONST_B01_FProductionDate As Long = 7 '原价（含税）
Private Const CONST_B01_FShelfLife As Long = 8 '实际销售金额（含税）
Private Const CONST_B01_FExpiryDate As Long = 9 '原销售金额（含税）=原价（含税）*数量
Private Const CONST_B01_FBoxBarcode As Long = 10 ' 折扣金额（不含税）
Private Const CONST_B01_FQty As Long = 11 ' 折扣金额（含税）
Private Const CONST_B01_FSPID As Long = 12 ' 折扣率
Private Const CONST_B01_FReleaseDate As Long = 13 ' 不含税销售金额
Private Const CONST_B01_FCustomerCode As Long = 14 ' 税额
Private Const CONST_B01_FCustomerName As Long = 15 ' 税率
Private Const CONST_B01_FAddress As Long = 16 ' 收银员
Private Const CONST_B01_FContact As Long = 17 ' 收银员
Private Const CONST_B01_FPhone As Long = 18 ' 收银员


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
' Procedure : SynBillB01
' DateTime  : 2013-1-25 00:05
' Author    :
' Purpose   : 生成销售出库单并审核 (POS TO ERP)
'---------------------------------------------------------------------------------------
'
Public Function SynBillB01(ByVal sDsn As String, varDatas As Variant, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String) As Boolean
    Dim CNN As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim clsBill As AutoOldBill
    Dim bRet As Boolean
    Dim lRow As Long
    Dim i As Long
    Dim dblDiscountAmount As Double '折扣额
    Dim dblMarkDownAmt As Double ''汇总折扣代码为 PRWN 的金额
    Dim dblAmt As Double '金额
    Dim dblDiscount As Double

    Dim lFItemID As Long
    Dim lUnitID As Long
    Dim dRate As Double

    Dim lSPGroupID As Long
    Dim blnIsSPGrop As Boolean
    Dim blnCheck As Boolean
    Dim sSourceBillNo As String
    Dim lROB As Long
    Dim blnEffective As Boolean ''是否存在有效物料
    Dim isUpdateStore As Boolean

    Dim strMarkDown As String

    '批号和保质期管理
    Dim lStockID As Long
    Dim lSPID As Long
    Dim strLotNo As String
    Dim strKFDate As String
    Dim strKFPeriod As String
    
    Dim strSQL As String
    Dim StrSqlTemp As String
    
    
    Dim objTypeLib As Object
    Dim strUUID As String
    Dim lSOInterID As Long
    
    Dim strBatch As String
    Dim strPeriodDate As String
    Dim strProduct As String
    Dim QTY As Long
    Dim ROW As Long
    
    Dim bAutoInv As Boolean
    
On Error GoTo SynBillB01_Error

    InitDataEnv sDsn
    Set CNN = datasource.Connection

    Set clsBill = New AutoOldBill

    lRow = 1
    blnCheck = False

    blnEffective = False
'    GetObjectContext.SetAbort

    lROB = 1
    

    '使用GUID作为一次事务的标识
    Set objTypeLib = CreateObject("Scriptlet.TypeLib")
    strUUID = CStr(objTypeLib.Guid)
    strUUID = Mid(strUUID, 1, InStr(1, strUUID, "}"))
    

    StrSqlTemp = ""
    
    With clsBill
        .sDsn = sDsn

        bRet = .NewBill(21)

        If bRet = True Then
            For i = LBound(varDatas) To UBound(varDatas)
            
'                If UBound(varDatas(i)) <> 15 Then
'                    sErr = "Data format error, please contact system administrator for checking."
'                    Exit Function
'                End If

                If i = 0 Then '表头
                
                    '检查POS Number是否存在，存在则不允许重复导入
                    Set rs = CNN.Execute("select 1 from ICStockBill where FTranType=21 and isnull(FPOSNumber,'')='" & CStr(varDatas(i)(1)) & "' and FROB=" & CStr(varDatas(i)(2)))
                    If rs.RecordCount > 0 Then
                        sErr = "POS Number[" & CStr(varDatas(i)(1)) & "] has existed in ERP system."
                        GoTo SynBillB01_Error
                    End If
                                
                    .SetHeadValue "FDate", Format$(CDate(varDatas(i)(3)), "yyyy-mm-dd") '日期
                    .SetHeadValue "FExplanation", Format(varDatas(i)(3), "HH:mm:ss") '日期
                    .SetHeadValue "FPOSNumber", CStr(varDatas(i)(1))  'POS Number
                    
                    .SetHeadValue "FConsigneeAdd", CStr(varDatas(i)(15))  'Payment Method
                    
                    .SetHeadValue "FSaleStyle", 101 '销售方式:默认赊销
                    .SetHeadValue "FMarketingStyle", 12530 '销售业务类型:默认销售出库
                    
                    .SetHeadValue "FROB", Val(varDatas(i)(2)) '红蓝字标识
                    lROB = Val(varDatas(i)(2))
                                    
                    '检查客户是否存在
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 1)
                    If lFItemID = 0 Then
                        sErr = "Customer Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillB01_Error
                    End If
                    .SetHeadValue "FSupplyID", lFItemID
                    '判断客户是否属于自动生成发票的客户
                    Set rs = CNN.Execute("select isnull(FAutoInv,0) FAutoInv from t_organization where FItemID=" & lFItemID)
                    bAutoInv = rs.Fields("FAutoInv").Value

                
                    '检查部门是否存在
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 2)
                    If lFItemID = 0 Then
                        sErr = "Department Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillB01_Error
                    End If
                    .SetHeadValue "FDeptID", lFItemID
                    
                    '检查职员是否存在
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 3)
                    If lFItemID = 0 Then
                        sErr = "Employee Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillB01_Error
                    End If
                    .SetHeadValue "FFManagerID", lFItemID
                    .SetHeadValue "FSManagerID", lFItemID
                
                End If
                
                
                '表体赋值

                .AddRow lRow
                
                ROW = lRow
                
                
                 '检查物料是否存在
                 lFItemID = GetItemIDByNumber(CNN, CStr(varDatas(i)(5)), 4, lUnitID)
                 If lFItemID = 0 Then
                     sErr = "Material Master Data [" & CStr(varDatas(i)(5)) & "] is NOT created in ERP system."
                     GoTo SynBillB01_Error
                 End If
                .SetEntryValue "FItemID", lFItemID, lRow            '物料
                .SetEntryValue "FUnitID", lUnitID, lRow            '单位
                
                '数量必须大于0
                If Val(varDatas(i)(6)) <= 0 Then
                    sErr = "The quantity can NOT be zero."
                    GoTo SynBillB01_Error
                End If
                .SetEntryValue "Fauxqty", Val(varDatas(i)(6)) * lROB, lRow           '数量
                .SetEntryValue "FAuxQtyMust", Val(varDatas(i)(6)) * lROB, lRow          '应发数量
                
                '检查仓库是否存在
                lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 5)
                If lFItemID = 0 Then
                     sErr = "Warehouse Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                     GoTo SynBillB01_Error
                 End If
                .SetEntryValue "FDCStockID", lFItemID, lRow        '发货仓库
            
                .SetEntryValue "FConsignPrice", CNulls(Val(varDatas(i)(7)), 0), lRow           ''单价
                .SetEntryValue "FDiscountAmount", CNulls(Val(varDatas(i)(12)), 0) * lROB, lRow          ''折扣额
                .SetEntryValue "FDiscountRate", (1 - CNulls(Val(varDatas(i)(11)), 0)) * 100, lRow       ''折扣率
                .SetEntryValue "FConsignAmount", CNulls(Val(varDatas(i)(13)), 0) * lROB, lRow          ''销售金额
                
                
                .SetEntryValue "FNote", CStr(varDatas(i)(14)), lRow          ''备注
                
                lRow = lRow + 1
NextRow:
'                blnCheck = True ''是否存在有效物料
            Next
            
            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
'                checkBillData sDsn, lInterID, 21, 0, sErr, 16394 '审核单据
                
                If checkBillData(sDsn, lInterID, 21, 0, sErr, 16394) = True Then  '审核单据
                    
                    '自动生成关联发票
'                    If bAutoInv = True Then
                        If GenerateSI(sDsn, lInterID, 16394, sErr) = False Then
                            GoTo SynBillB01_Error
                        End If
'                    End If
                    
                Else
                    sErr = "Approve the sales delivery failed: " & sErr
                    GoTo SynBillB01_Error
                End If
                
            End If
        End If
    End With

    Set clsBill = Nothing
    Set CNN = Nothing
    SynBillB01 = True
    GetObjectContext.SetComplete

   Exit Function

SynBillB01_Error:
    SynBillB01 = False
    sErr = "Generate Sales Delivery Failed：" & sErr & Err.Description
    GetObjectContext.SetAbort
    Set clsBill = Nothing
    Set CNN = Nothing
    

End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateSI
' DateTime  : 2018-9-5
' Author    :
' Purpose   : 如果客户基础资料勾选了自动生成发票，生成销售发票并审核
'---------------------------------------------------------------------------------------
'
Private Function GenerateSI(ByVal sDsn As String, ByVal lInterID As Long, ByVal lUserID As Long, ByRef sErr As String) As Boolean
 
    Dim CNN As ADODB.Connection
    Dim rs As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim clsBill As AutoOldBill
    Dim bRet As Boolean
    Dim lRow As Long
    Dim i, j As Long
    Dim lFItemID As Long
    Dim lUnitID As Long
    Dim strSQL As String
    Dim strBillNo As String
    
    Dim lROB As Long, ROW As Long
    
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
            
            .SetHeadValue "FItemClassID", 1
            .SetHeadValue "FClassTypeID", 1000002

            strSQL = "select FAccountID from t_Account where FNumber='1122.01'"
            Set rsTemp = ExecSQL(strSQL, sDsn)
            .SetHeadValue "FCussentAcctID", rsTemp.Fields("FAccountID").Value    '往来科目

        
            '表体赋值
            For j = 1 To rs.RecordCount
                
                .AddRow lRow
                
                ROW = lRow

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
'                .SetEntryValue "FOrderInterID", rs.Fields("FInterID").Value, lRow
'                .SetEntryValue "FOrderEntryID", rs.Fields("FEntryID").Value, lRow
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


Private Function ExecSQL(ByVal ssql As String, ByVal dsn As String) As ADODB.Recordset
    Dim OBJ As Object
    Dim rs As ADODB.Recordset

    Set OBJ = CreateObject("BillDataAccess.GetData")
    Set rs = OBJ.ExecuteSQL(dsn, ssql)
    Set OBJ = Nothing
    Set ExecSQL = rs
    Set rs = Nothing
End Function






