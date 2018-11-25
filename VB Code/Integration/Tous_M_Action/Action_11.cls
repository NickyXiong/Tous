VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "Action_11"
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
Public Function SynBillB01(ByVal sDsn As String, varDatas As Variant, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String, ByRef strUUID As String) As Boolean
    Dim cnn As ADODB.Connection
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

    Dim strBarCode As String
    Dim strMarkDown As String

    '批号和保质期管理
    Dim lStockID As Long
    Dim lSPID As Long
    Dim strLotNo As String
    Dim strKFDate As String
    Dim strKFPeriod As String
    
    Dim StrSql As String
    Dim StrSqlTemp As String
    
    Dim strBatch As String
    Dim strPeriodDate As String
    Dim strProduct As String
    Dim Row As Long
    Dim QTY As Long
    Dim EntryID As Long
    
    Dim objTypeLib As Object
'    Dim strUUID As String
    Dim lSOInterID As Long
    
    Dim strDSN2 As String
    Dim rsYearMonth As ADODB.Recordset
    Dim ssql As String

On Error GoTo SynBillB01_Error

    InitDataEnv sDsn
    Set cnn = datasource.Connection

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
            
                If UBound(varDatas(i)) <> 18 Then
                    sErr = "导入文件格式错误，请联系系统管理员"
                    Exit Function
                End If

                If i = 0 Then '表头

                    sSourceBillNo = Trim(CStr(varDatas(i)(CONST_B01_FBillNo)))

                    If IsExitBill(cnn, sSourceBillNo, "SEOrder", 81, lSOInterID, "FBillNo") = False Then
                        sErr = "销售订单 [" & sSourceBillNo & "] 不存在"
                        Exit Function
                    Else
                        StrSql = "select t3.FBarcode,t1.FCustID,t1.FDeptID,t1.FEmpID,t1.FMangerID,t2.FEntryID,t2.FItemID,t2.FAuxTaxPrice,FTaxAmount FDiscountAmt,t2.FTaxRate FDiscountRate,FAuxStockQty,(case when day(t1.FHeadSelfS0154)>25 then 102 else 101 end) FSaleStyle,t1.FHeadSelfS0156 SalesMode"
                        StrSql = StrSql & vbCrLf & "from SEOrder t1 inner join SEOrderEntry t2 on t1.FInterID=t2.FInterID"
                        StrSql = StrSql & vbCrLf & "inner join t_icitem t3 on t2.FItemID=t3.FItemID"
                        StrSql = StrSql & vbCrLf & "where t1.FInterID =" & lSOInterID
                        Set rsTemp = cnn.Execute(StrSql)
                    End If
                
                    .SetHeadValue "FMarketingStyle", 12530 '销售业务类型 ：默认销售出库类型
                    
                    If Len(CStr(varDatas(i)(CONST_B01_FDATE))) <= 0 Then
                        sErr = "销售订单[" & sSourceBillNo & "]出库信息:日期不能为空"
                        Exit Function
                    End If
                    .SetHeadValue "FHeadSelfB0157", rsTemp.Fields("SalesMode").Value   '销售类型
                    .SetHeadValue "FDate", Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd")  ' 日期
                    .SetHeadValue "FSettleDate", Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd")  ' 收款日期


                    .SetHeadValue "FDeptID", rsTemp.Fields("FDeptID").Value   '部门
                    .SetHeadValue "FSupplyID", rsTemp.Fields("FCustID").Value  '购货单位

                    '销售方式： 赊销
                    .SetHeadValue "FSaleStyle", 101


                    '职员字段全部默认为 001
                    lFItemID = GetItemID(cnn, "001", 3)
                    If lFItemID = 0 Then
                        sErr = "Can NOT find Employee [001]"
                        Exit Function
                    End If
                    .SetHeadValue "FFManagerID", lFItemID '发货必录
                    .SetHeadValue "FEmpID", rsTemp.Fields("FEmpID").Value   '业务员"
                    .SetHeadValue "FSManagerID", lFItemID   '保管
                    .SetHeadValue "FManagerID", lFItemID  '主管


                    .SetHeadValue "FRob", lROB '红蓝字
                End If
                
                
                rsTemp.Filter = "FEntryID=" & Trim(CStr(varDatas(i)(CONST_B01_FEntryID)))
                If Not rsTemp Is Nothing Then
                    If rsTemp.RecordCount > 0 Then
                    Else
'                        sErr = "条目号 [" & sSourceBillNo & "] 不存在"
                        sErr = "销售订单[" & sSourceBillNo & "]出库信息:条目号不存在"
                        Exit Function
                    End If
                End If

                '表体赋值

                .AddRow lRow
                
                '相同物料批号到期日的数量合并 modify 20160316
                If i <> 0 And strProduct = Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) And strBatch = CStr(Trim(varDatas(i)(CONST_B01_FLotNumber))) _
                And strPeriodDate = Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate))) And EntryID = Trim(CStr(varDatas(i)(CONST_B01_FEntryID))) Then
                    QTY = QTY + Val(varDatas(i)(CONST_B01_FQty)) * lROB
                    .SetEntryValue "Fauxqty", QTY, Row         '数量
                    .SetEntryValue "FAuxQtyMust", (QTY - Val(rsTemp.Fields("FAuxStockQty").Value)) * lROB, Row
                    
                    StrSqlTemp = StrSqlTemp & vbCrLf & "insert T_t_Sales  values('销售订单','" & sSourceBillNo & _
                        "'," & lStockID & "," & lSPID & ",'" & Trim(CStr(varDatas(i)(CONST_B01_FBoxBarcode))) & _
                        "',1,'" & Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd") & _
                        "','" & Format(varDatas(i)(CONST_B01_FDATE), "yyyymmddhhmmss") & _
                        "',0,0,'" & strUUID & "',1," & lStockID & "," & lSPID & ",'" & Format(Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate))), "yyyy-mm") & "')"
                    GoTo NextRow
                Else
                    Row = Row + 1
                    QTY = 0
                End If
                Row = lRow

                lFItemID = GetItemID(cnn, "C014", 5)
                If lFItemID = 0 Then
'                    sErr = "仓库[C004]不存在"
                    sErr = "销售订单[" & sSourceBillNo & "]出库信息:仓库[C014]不存在"
                    Exit Function
                End If
                .SetEntryValue "FDCStockID", lFItemID, lRow            '发货仓库
                lStockID = lFItemID
                
                EntryID = Trim(CStr(varDatas(i)(CONST_B01_FEntryID)))
                
                blnIsSPGrop = GetIsDCSP(cnn, lStockID, lSPGroupID)
                If blnIsSPGrop = True Then
                    '仓位需根据仓库属性仓位组查找仓位信息
                    lFItemID = GetDCSPID(cnn, Trim(CStr(varDatas(i)(CONST_B01_FSPID))), lSPGroupID)
                    If lFItemID = 0 Then
'                        sErr = "仓位[" & Trim(CStr(varDatas(i)(CONST_B01_FSPID))) & "] 不存在"
                        sErr = "销售订单[" & sSourceBillNo & "]出库信息:仓位[" & Trim(CStr(varDatas(i)(CONST_B01_FSPID))) & "] 不存在"
                        Exit Function
                    End If
                    lSPID = lFItemID
                    .SetEntryValue "FDCSPID", lFItemID, lRow              '仓位
                End If

                lFItemID = GetItemID(cnn, Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))), 4, lUnitID, dRate)
                If lFItemID = 0 Then
'                    sErr = "发出商品[" & Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) & "]不存在"
                    sErr = "销售订单[" & sSourceBillNo & "]出库信息:发出商品[" & Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) & "]不存在"
                    Exit Function
                End If
                
                If Val(rsTemp.Fields("FItemID").Value) <> lFItemID Then
'                    sErr = "发出商品[" & Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) & "]与销售订单的商品不一致，请检查数据或条目号是否正确"
                    sErr = "销售订单[" & sSourceBillNo & "]出库信息:发出商品[" & Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) & "]与销售订单的商品不一致，请检查数据或条目号是否正确"
                    Exit Function
                End If
                
                .SetEntryValue "FItemID", lFItemID, lRow     '物料
                .SetEntryValue "FUnitID", lUnitID, lRow      '单位
                
                strProduct = Trim(CStr(varDatas(i)(CONST_B01_FItemNumber)))
                strBarCode = Trim(CStr(varDatas(i)(CONST_B01_FBoxBarcode)))

                '批号和保质期
                If Len(CStr(varDatas(i)(CONST_B01_FLotNumber))) <= 0 Then
'                    sErr = "销售订单 [" & sSourceBillNo & "] :批号不能为空"
                    sErr = "销售订单[" & sSourceBillNo & "]出库信息:批号不能为空"
                    Exit Function
                End If
                .SetEntryValue "FBatchNo", CStr(Trim(varDatas(i)(CONST_B01_FLotNumber))), lRow              ''仓库
                strBatch = CStr(Trim(varDatas(i)(CONST_B01_FLotNumber)))
                
                If Len(CStr(varDatas(i)(CONST_B01_FProductionDate))) <= 0 Then
'                    sErr = "销售订单 [" & sSourceBillNo & "] :生产日期不能为空"
                    sErr = "销售订单[" & sSourceBillNo & "]出库信息:生产日期不能为空"
                    Exit Function
                End If
'                .SetEntryValue "FKFDate", CStr(Trim(varDatas(i)(CONST_B01_FProductionDate))), lRow              ''仓库
                .SetEntryValue "FKFDate", DateAdd("d", -Val(varDatas(i)(CONST_B01_FShelfLife)), Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate)))), lRow            ''生产日期
                
                If Len(CStr(varDatas(i)(CONST_B01_FShelfLife))) <= 0 Then
'                    sErr = "销售订单 [" & sSourceBillNo & "] :保质期不能为空"
                    sErr = "销售订单[" & sSourceBillNo & "]出库信息:保质期不能为空"
                    Exit Function
                End If
                .SetEntryValue "FKFPeriod", Val(varDatas(i)(CONST_B01_FShelfLife)), lRow              ''仓库
                
                If Len(CStr(varDatas(i)(CONST_B01_FExpiryDate))) <= 0 Then
'                    sErr = "销售订单 [" & sSourceBillNo & "] :到期日期不能为空"
                    sErr = "销售订单[" & sSourceBillNo & "]出库信息:到期日期不能为空"
                    Exit Function
                End If
                .SetEntryValue "FPeriodDate", Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate))), lRow '有效期至
                strPeriodDate = Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate)))

                If Val(CStr(varDatas(i)(CONST_B01_FQty))) <= 0 Then
'                    sErr = "销售订单 [" & sSourceBillNo & "] :出货数量必须为正数"
                    sErr = "销售订单[" & sSourceBillNo & "]出库信息:出货数量必须为正数"
                    Exit Function
                End If
                .SetEntryValue "Fauxqty", Val(varDatas(i)(CONST_B01_FQty)) * lROB, lRow           '数量
                .SetEntryValue "FAuxQtyMust", (Val(varDatas(i)(CONST_B01_FQty)) - Val(rsTemp.Fields("FAuxStockQty").Value)) * lROB, lRow       '应发数量
                QTY = Val(varDatas(i)(CONST_B01_FQty)) * lROB
                
                '单价
                .SetEntryValue "FConsignPrice", Val(rsTemp.Fields("FAuxTaxPrice").Value), lRow
                '销售金额
                .SetEntryValue "FConsignAmount", Val(varDatas(i)(CONST_B01_FQty)) * _
                                Val(rsTemp.Fields("FAuxTaxPrice").Value) * lROB * _
                                (1 - Val(rsTemp.Fields("FDiscountRate").Value) / 100), lRow
                
                '折扣率
                .SetEntryValue "FDiscountRate", Val(rsTemp.Fields("FDiscountRate").Value), lRow
                '含税折扣额
'                .SetEntryValue "FDiscountAmount", Val(varDatas(i)(CONST_B01_FQty)) * _
'                                Val(rsTemp.Fields("FAuxTaxPrice").Value) * lROB * _
'                                (Val(rsTemp.Fields("FDiscountRate").Value) / 100), lRow
                .SetEntryValue "FDiscountAmount", Val(rsTemp.Fields("FDiscountAmt").Value), lRow
                
                ''关联关系字段
                .SetEntryValue "FSourceTranType", 81, lRow           '
                .SetEntryValue "FSourceInterID", lSOInterID, lRow            '
                .SetEntryValue "FSourceBillNo", sSourceBillNo, lRow
                .SetEntryValue "FSourceEntryID", CNulls(rsTemp.Fields("FEntryID"), 0), lRow '
                .SetEntryValue "FOrderInterID", lSOInterID, lRow           '
                .SetEntryValue "FOrderEntryID", CNulls(rsTemp.Fields("FEntryID"), 0), lRow           '
                .SetEntryValue "FOrderBillNo", sSourceBillNo, lRow
                
                StrSqlTemp = StrSqlTemp & vbCrLf & "insert T_t_Sales  values('销售订单','" & sSourceBillNo & _
                                    "'," & lStockID & "," & lSPID & ",'" & Trim(CStr(varDatas(i)(CONST_B01_FBoxBarcode))) & _
                                    "',1,'" & Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd") & _
                                    "','" & Format(varDatas(i)(CONST_B01_FDATE), "yyyymmddhhmmss") & _
                                    "',0,0,'" & strUUID & "',1," & lStockID & "," & lSPID & ",'" & Format(Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate))), "yyyy-mm") & "')"
                
                lRow = lRow + 1
NextRow:
'                blnCheck = True ''是否存在有效物料
            Next
            
            cnn.Execute StrSqlTemp

            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
'                checkBillData sDsn, lInterID, 21, 0, sErr, 16394 '审核单据
                
                StrSqlTemp = "update T_t_Sales set FTransType=21,FInterID=" & lInterID & " where FUUID='" & strUUID & "'"
                cnn.Execute StrSqlTemp
            End If
        End If
    End With
    
'    '=====导入新数据库======
'
'     '获取新数据库连接字符串
'    strDSN2 = TransfersDsn("st20160407", sDsn)
'
'    '查询本次导入的数据按到期日年月分组导入
'    ssql = "select isnull(FYearMonth,'') FYearMonth from T_t_sales where FUUID='" & strUUID & "' and isnull(FYearMonth,'') <>'' group by FYearMonth"
'    Set rsYearMonth = ExecSQL(ssql, sDsn)
'
'
'    '根据到期日检查对应表在新数据库中是否存在,不存在则新建
'    For i = 1 To rsYearMonth.RecordCount
'        ssql = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[T_t_Sales_" & rsYearMonth("FYearMonth") & "]') AND type in (N'U'))"
'        ssql = ssql & vbCrLf & "CREATE TABLE [dbo].[T_t_Sales_" & rsYearMonth("FYearMonth") & "]("
'        ssql = ssql & vbCrLf & "[FType] [nvarchar](255) NULL,"
'        ssql = ssql & vbCrLf & "[FBillNo] [nvarchar](255) NULL,"
'        ssql = ssql & vbCrLf & "[FWareHouse] [int] NULL,"
'        ssql = ssql & vbCrLf & "[FSPID] [int] NULL,"
'        ssql = ssql & vbCrLf & "[FBarCode] [nvarchar](50) NULL,"
'        ssql = ssql & vbCrLf & "[FBarSgin] [nvarchar](50) NULL,"
'        ssql = ssql & vbCrLf & "[FDate] [nvarchar](50) NULL,"
'        ssql = ssql & vbCrLf & "[FImportTime] [nvarchar](50),"
'        ssql = ssql & vbCrLf & "[FTransType] [int] NULL,"
'        ssql = ssql & vbCrLf & "[FInterID] [int] NULL,"
'        ssql = ssql & vbCrLf & "[FUUID] [nvarchar](255) NULL,"
'        ssql = ssql & vbCrLf & "[FIsExported] int null default (0),"
'        ssql = ssql & vbCrLf & "[FDefaultWareHouse] [int] NULL,"
'        ssql = ssql & vbCrLf & "[FDefaultSPID] [int] NULL,"
'        ssql = ssql & vbCrLf & "[FYearMonth] [nvarchar](10) NULL"
'        ssql = ssql & vbCrLf & ") ON [PRIMARY]"
'        ExecSQL ssql, strDSN2
'
'        '将本次导入数据按到期年月分组导入对应新数据库表中
'        ssql = "insert into [T_t_sales_" & rsYearMonth.Fields("FYearMonth") & "] select * from [st20160414].dbo.T_t_sales where FUUID='" & strUUID & "' and FYearMonth='" & rsYearMonth.Fields("FYearMonth") & "'"
'        ExecSQL ssql, strDSN2
'
'    rsYearMonth.MoveNext
'    Next

    
    Set clsBill = Nothing
    Set cnn = Nothing
    SynBillB01 = True
    GetObjectContext.SetComplete

   Exit Function

SynBillB01_Error:
    SynBillB01 = False
    GetObjectContext.SetAbort
    Set clsBill = Nothing
    Set cnn = Nothing
    sErr = "生成销售出库单失败：" & sErr & Err.Description
    

End Function


'---------------------------------------------------------------------------------------
' Procedure : SynBillD01
' DateTime  : 2013-1-25 00:05
' Author    :
' Purpose   : 生成调拨单并审核
'---------------------------------------------------------------------------------------
'
Public Function SynBillD01(ByVal sDsn As String, varDatas As Variant, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String) As Boolean
    Dim cnn As ADODB.Connection
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
    Dim lDCStockID As Long
    Dim lSPID As Long
    Dim strLotNo As String
    Dim strKFDate As String
    Dim strKFPeriod As String
    
    Dim StrSql As String
    Dim StrSqlTemp As String
    
    
    Dim objTypeLib As Object
    Dim strUUID As String
    Dim lSOInterID As Long

On Error GoTo SynBillD01_Error

    InitDataEnv sDsn
    Set cnn = datasource.Connection

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

        bRet = .NewBill(41)

        If bRet = True Then
            For i = LBound(varDatas) To UBound(varDatas)
            


                If i = 0 Then '表头

                    sSourceBillNo = Trim(CStr(varDatas(i)(CONST_B01_FBillNo)))

                    If IsExitBill(cnn, sSourceBillNo, "ICSTOCKBILL", 41, , "FBillNo") = True Then
                        sErr = "调拨单[" & sSourceBillNo & "]已生成，不允许重复导入"
                        Exit Function
                    End If
                    .SetHeadValue "FBillNo", sSourceBillNo  ' 单号

                    If Len(CStr(varDatas(i)(CONST_B01_FDATE))) <= 0 Then
                        sErr = "调拨单[" & sSourceBillNo & "] :日期不能为空"
                        Exit Function
                    End If
                    .SetHeadValue "FDate", Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd")  ' 日期
'                    .SetHeadValue "FSettleDate", Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd")  ' 收款日期


                    '职员字段全部默认为 001
                    lFItemID = GetItemID(cnn, "001", 3)
                    If lFItemID = 0 Then
                        sErr = "职员[001]不存在"
                        Exit Function
                    End If
                    .SetHeadValue "FFManagerID", lFItemID '发货必录
                    .SetHeadValue "FEmpID", lFItemID   '业务员"
                    .SetHeadValue "FSManagerID", lFItemID   '保管
                    .SetHeadValue "FManagerID", lFItemID  '主管

                    .SetHeadValue "FRob", lROB '红蓝字

                End If
                

                '表体赋值

                .AddRow lRow
                
                
                
                lFItemID = GetItemID(cnn, "C014", 5)
                If lFItemID = 0 Then
                    sErr = "调拨单[" & sSourceBillNo & "] :发货仓库[C014] 不存在"
                    Exit Function
                End If
                .SetEntryValue "FSCStockID", lFItemID, lRow            '发货仓库
                lStockID = lFItemID
                
                
                blnIsSPGrop = GetIsDCSP(cnn, lStockID, lSPGroupID)
                If blnIsSPGrop = True Then
                    '仓位需根据仓库属性仓位组查找仓位信息
                    lFItemID = GetDCSPID(cnn, Trim(CStr(varDatas(i)(CONST_B01_FSPID))), lSPGroupID)
                    If lFItemID = 0 Then
                        sErr = "调拨单[" & sSourceBillNo & "] :发货仓位[" & Trim(CStr(varDatas(i)(CONST_B01_FSPID))) & "] 不存在"
                        Exit Function
                    End If
                    .SetEntryValue "FSCSPID", lFItemID, lRow              '仓位
                    
'                    lFItemID = GetDCSPID(CNN, Trim(CStr(varDatas(i)(CONST_D01_FDestBinID))), lSPGroupID)
'                    If lFItemID = 0 Then
'                        sErr = "收货仓位[" & Trim(CStr(varDatas(i)(CONST_D01_FDestBinID))) & "] 不存在"
'                        Exit Function
'                    End If
'                    lSPID = lFItemID
'                    .SetEntryValue "FDCSPID", lFItemID, lRow              '仓位
                End If
                
                lFItemID = GetItemID(cnn, Trim(CStr(varDatas(i)(CONST_B01_FCustomerCode))), 5)
                If lFItemID = 0 Then
                    sErr = "调拨单[" & sSourceBillNo & "] :收货仓库[" & Trim(CStr(varDatas(i)(CONST_B01_FCustomerCode))) & "] 不存在"
                    Exit Function
                End If
                .SetEntryValue "FDCStockID", lFItemID, lRow            '收货仓库
                lDCStockID = lFItemID
                
                
                lFItemID = GetItemID(cnn, Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))), 4, lUnitID, dRate)
                If lFItemID = 0 Then
                    sErr = "调拨单[" & sSourceBillNo & "] :商品[" & Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) & "]不存在"
                    Exit Function
                End If
                .SetEntryValue "FItemID", lFItemID, lRow     '物料
                .SetEntryValue "FUnitID", lUnitID, lRow      '单位
                
                
                '批号和保质期
                If Len(CStr(varDatas(i)(CONST_B01_FLotNumber))) <= 0 Then
                    sErr = "调拨单[" & sSourceBillNo & "] :批号不能为空"
                    Exit Function
                End If
                .SetEntryValue "FBatchNo", CStr(Trim(varDatas(i)(CONST_B01_FLotNumber))), lRow              ''仓库
                
                If Len(CStr(varDatas(i)(CONST_B01_FProductionDate))) <= 0 Then
                    sErr = "调拨单[" & sSourceBillNo & "] :生产日期不能为空"
                    Exit Function
                End If
                .SetEntryValue "FKFDate", CStr(Trim(varDatas(i)(CONST_B01_FProductionDate))), lRow              ''仓库
                
                If Len(CStr(varDatas(i)(CONST_B01_FShelfLife))) <= 0 Then
                    sErr = "调拨单[" & sSourceBillNo & "] :保质期不能为空"
                    Exit Function
                End If
                .SetEntryValue "FKFPeriod", Val(varDatas(i)(CONST_B01_FShelfLife)), lRow              ''仓库
                
                If Len(CStr(varDatas(i)(CONST_B01_FExpiryDate))) <= 0 Then
                    sErr = "调拨单[" & sSourceBillNo & "] :到期日期不能为空"
                    Exit Function
                End If
                .SetEntryValue "FPeriodDate", Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate))), lRow '有效期至


                If Val(CStr(varDatas(i)(CONST_B01_FQty))) <= 0 Then
                    sErr = "调拨单[" & sSourceBillNo & "] :移库数量必须为正数"
                    Exit Function
                End If
                .SetEntryValue "Fauxqty", Val(varDatas(i)(CONST_B01_FQty)), lRow           '数量
                
                
                StrSqlTemp = StrSqlTemp & vbCrLf & "insert T_t_Sales  values('调拨单','" & sSourceBillNo & _
                                    "'," & lDCStockID & ",0,'" & Trim(CStr(varDatas(i)(CONST_B01_FBoxBarcode))) & _
                                    "',1,'" & Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd") & _
                                    "','" & Format(varDatas(i)(CONST_B01_FDATE), "yyyymmddhhmmss") & _
                                    "',0,0,'" & strUUID & "',1," & lStockID & "," & lSPID & ")"
                
                lRow = lRow + 1

'                blnCheck = True ''是否存在有效物料
            Next
            
            cnn.Execute StrSqlTemp

            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
'                checkBillData sDsn, lInterID, 41, 0, sErr, 16394 '审核单据
                
                StrSqlTemp = "update T_t_Sales set FTransType=41,FInterID=" & lInterID & " where FUUID='" & strUUID & "'"
                cnn.Execute StrSqlTemp
            End If
        End If
    End With

    Set clsBill = Nothing
    Set cnn = Nothing
    SynBillD01 = True
    GetObjectContext.SetComplete

   Exit Function

SynBillD01_Error:
    SynBillD01 = False
    GetObjectContext.SetAbort
    Set clsBill = Nothing
    Set cnn = Nothing
    sErr = "生成调拨单失败：" & sErr

End Function

Public Function ExecSQL(ByVal ssql As String, ByVal dsn As String) As ADODB.Recordset
    Dim OBJ As Object
    Dim rs As ADODB.Recordset

    Set OBJ = CreateObject("BillDataAccess.GetData")
    Set rs = OBJ.ExecuteSQL(dsn, ssql)
    Set OBJ = Nothing
    Set ExecSQL = rs
    Set rs = Nothing
End Function



