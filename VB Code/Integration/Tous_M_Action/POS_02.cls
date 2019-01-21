VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************
'Function: 生成外购入库单并审核 (POS TO ERP)
'Date:
'*******************************************************

Option Explicit

Private datasource As CDataSource
Private m_oParse As CParse
Private strLanguage As String


'Private Const CONST_B01_FBillNo As Long = 0 '销售店铺
'
''Added by Nicky  - 20150504
''增加是否销售类型字段
'Private Const CONST_B01_FType As Long = 1 '销售出库单代码
''Added by Nicky  - 20150504
'
'Private Const CONST_B01_FDATE As Long = 2 '销售出库单代码
'Private Const CONST_B01_FEntryID  As Long = 3 '销售/退货
'Private Const CONST_B01_FItemNumber  As Long = 4 '日期
'Private Const CONST_B01_FUOM As Long = 5 '物料代码
'Private Const CONST_B01_FLotNumber As Long = 6 ' 数量
'Private Const CONST_B01_FProductionDate As Long = 7 '原价（含税）
'Private Const CONST_B01_FShelfLife As Long = 8 '实际销售金额（含税）
'Private Const CONST_B01_FExpiryDate As Long = 9 '原销售金额（含税）=原价（含税）*数量
'Private Const CONST_B01_FBoxBarcode As Long = 10 ' 折扣金额（不含税）
'Private Const CONST_B01_FQty As Long = 11 ' 折扣金额（含税）
'Private Const CONST_B01_FSPID As Long = 12 ' 折扣率
'Private Const CONST_B01_FReleaseDate As Long = 13 ' 不含税销售金额
'Private Const CONST_B01_FCustomerCode As Long = 14 ' 税额
'Private Const CONST_B01_FCustomerName As Long = 15 ' 税率
'Private Const CONST_B01_FAddress As Long = 16 ' 收银员
'Private Const CONST_B01_FContact As Long = 17 ' 收银员
'Private Const CONST_B01_FPhone As Long = 18 ' 收银员


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
' Procedure : SynBillA01
' DateTime  :
' Author    :
' Purpose   : 生成外购入库单并审核 (POS TO ERP)
'---------------------------------------------------------------------------------------
'
Public Function SynBillA01(ByVal sDsn As String, varDatas As Variant, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String) As Boolean
    Dim CNN As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim rsPOInfo As ADODB.Recordset
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
    
On Error GoTo SynBillA01_Error

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

        bRet = .NewBill(1)

        If bRet = True Then
            For i = LBound(varDatas) To UBound(varDatas)
            
'                If UBound(varDatas(i)) <> 15 Then
'                    sErr = "Data format error, please contact system administrator for checking."
'                    Exit Function
'                End If

                If i = 0 Then '表头
                
                    '检查POS Number是否存在，存在则不允许重复导入
                    Set rs = CNN.Execute("select 1 from ICStockBill where FTranType=1 and isnull(FPOSNumber,'')='" & CStr(varDatas(i)(1)) & "'")
                    If rs.RecordCount > 0 Then
                        sErr = "POS Number[" & CStr(varDatas(i)(1)) & "] has existed in ERP system."
                        GoTo SynBillA01_Error
                    End If
                    .SetHeadValue "FPOSNumber", CStr(varDatas(i)(1))  '日期
                                
                             
                    '检查采购订单是否存在
                    Set rsPOInfo = CNN.Execute("select t1.*,t2.*,t3.FNumber FItemNumber from POOrder t1 inner join POOrderEntry t2 on t1.FInterID=t2.FInterID inner join t_ICItem t3 on t2.FItemID=t3.FItemID where t1.FBillNo='" & CStr(varDatas(i)(2)) & "'")
                    If rsPOInfo.RecordCount <= 0 Then
                        sErr = "PO[" & CStr(varDatas(i)(2)) & "] does NOT exist in ERP system."
                        GoTo SynBillA01_Error
                    End If
                                                                
                    .SetHeadValue "FDate", CStr(varDatas(i)(4))  '日期
                    .SetHeadValue "FSupplyID", rsPOInfo.Fields("FSupplyID").Value '供应商默认取采购订单供应商
                    .SetHeadValue "FPOMode", 36680 '采购模式:  默认为普通采购
                    .SetHeadValue "FPOStyle", 252  '采购方式: 默认为赊购
                    
                    .SetHeadValue "FHQNumber", rsPOInfo.Fields("FHQNumber")
                    
                    .SetHeadValue "FEmpID", rsPOInfo.Fields("FEmpID")
                    .SetHeadValue "FFManagerID", rsPOInfo.Fields("FEmpID")
                    .SetHeadValue "FManagerID", rsPOInfo.Fields("FEmpID")
                    .SetHeadValue "FSManagerID", rsPOInfo.Fields("FEmpID")
                
                    .SetHeadValue "FDeptID", rsPOInfo.Fields("FDeptID")
                
                End If
                
                
                '表体赋值

                .AddRow lRow
                
                ROW = lRow
                
                rsPOInfo.Filter = "FItemNumber='" & CStr(varDatas(i)(5)) & "'"
                
                '检查物料是否存在
                 If rsPOInfo.RecordCount = 0 Then
                     sErr = "Material Master Data [" & CStr(varDatas(i)(5)) & "] does NOT exist in the original PO."
                     GoTo SynBillA01_Error
                 End If
                .SetEntryValue "FItemID", rsPOInfo.Fields("FItemID").Value, lRow            '物料
                .SetEntryValue "FUnitID", rsPOInfo.Fields("FUnitID").Value, lRow            '单位
                    
                '数量必须大于0
                If Val(varDatas(i)(6)) <= 0 Then
                    sErr = "The quantity can NOT be zero."
                    GoTo SynBillA01_Error
                End If
                .SetEntryValue "Fauxqty", Val(varDatas(i)(6)), lRow            '数量
                .SetEntryValue "FAuxQtyMust", Val(varDatas(i)(6)), lRow           '应收数量
                    
                '检查收货仓库
                lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 5)
                If lFItemID = 0 Then
                     sErr = "Warehouse Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                     GoTo SynBillA01_Error
                 End If
                .SetEntryValue "FDCStockID", lFItemID, lRow        '收货仓库
                    
'                     If Val(rsBillInfo.Fields("FIsStockMgr").Value) <> 0 And Val(rsBillInfo.Fields("FBinID").Value) = 0 Then
'                         sErr = "Bin Master Data [" & rsBillInfo.Fields("FBinCode").Value & "] is NOT created in ERP system or does not in the right bin group."
'                         GoTo SynBillA01_Error
'                     End If
'                    .SetEntryValue "FDCSPID", rsBillInfo.Fields("FBinID").Value, lRow        '收货仓位

                 '=========源单信息===========
                 .SetEntryValue "FSourceTranType", 71, lRow          ''源单类型
                 .SetEntryValue "FSourceBillNo", rsPOInfo.Fields("FBillNo").Value, lRow          ''源单编号
                 .SetEntryValue "FSourceInterId", rsPOInfo.Fields("FInterID").Value, lRow          ''源单内码
                 .SetEntryValue "FSourceEntryID", rsPOInfo.Fields("FEntryID").Value, lRow          ''源单行号
                 .SetEntryValue "FOrderInterID", rsPOInfo.Fields("FInterID").Value, lRow            '
                 .SetEntryValue "FOrderEntryID", rsPOInfo.Fields("FEntryID").Value, lRow           '
                 .SetEntryValue "FOrderBillNo", rsPOInfo.Fields("FBillNo").Value, lRow
                '=============================
                    
                .SetEntryValue "FNote", CStr(varDatas(i)(8)), lRow          ''备注
                
                lRow = lRow + 1
                
                rsPOInfo.Filter = ""
NextRow:
'                blnCheck = True ''是否存在有效物料
            Next
            
            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
                checkBillData sDsn, lInterID, 1, 0, sErr, 16394 '审核单据
                
            End If
        End If
    End With

    Set clsBill = Nothing
    Set CNN = Nothing
    SynBillA01 = True
    GetObjectContext.SetComplete

   Exit Function

SynBillA01_Error:
    SynBillA01 = False
    sErr = "Generate ：" & sErr & Err.Description
    GetObjectContext.SetAbort
    Set clsBill = Nothing
    Set CNN = Nothing
    

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


'---------------------------------------------------------------------------------------
' Procedure : SyncStockAdjustment
' DateTime  :
' Author    :
' Purpose   : 生成其他入库单并审核 (POS TO ERP)
'---------------------------------------------------------------------------------------
'
Public Function SyncStockAdjustment(ByVal sDsn As String, varDatas As Variant, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String) As Boolean
    Dim CNN As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim rsPOInfo As ADODB.Recordset
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
    Dim lTranType As Long
    

    
On Error GoTo SynBillA97_Error

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
        
        bRet = .NewBill(10)
        lTranType = 10

        If bRet = True Then
            For i = LBound(varDatas) To UBound(varDatas)

                If i = 0 Then '表头
                
                    '检查采购订单是否存在
                    Set rsPOInfo = CNN.Execute("select t1.*,t2.*,t3.FNumber FItemNumber from POOrder t1 inner join POOrderEntry t2 on t1.FInterID=t2.FInterID inner join t_ICItem t3 on t2.FItemID=t3.FItemID where t1.FBillNo='" & CStr(varDatas(i)(2)) & "'")
                    If rsPOInfo.RecordCount <= 0 Then
                        sErr = "PO[" & CStr(varDatas(i)(2)) & "] does NOT exist in ERP system."
                        GoTo SynBillA97_Error
                    End If
                    
                    '检查POS Number是否存在，存在则不允许重复导入
                    Set rs = ExecSQL("select 1 from ICStockBill where FTranType=" & lTranType & " and isnull(FPOSNumber,'')='" & CStr(varDatas(i)(1)) & "'", sDsn)
                    If rs.RecordCount > 0 Then
                        sErr = "POS Number[" & CStr(varDatas(i)(1)) & "] has existed in ERP system."
                        GoTo SynBillA97_Error
                    End If
                    .SetHeadValue "FPOSNumber", CStr(varDatas(i)(1))  '日期
                                
                    .SetHeadValue "FDate", CStr(varDatas(i)(4))  '日期
                    
                    
                    '检查部门是否存在
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 2)
                    If lFItemID = 0 Then
                        sErr = "Department Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillA97_Error
                    End If
                    .SetHeadValue "FDeptID", lFItemID
                    
                    
                    '检查职员是否存在
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 3)
                    If lFItemID = 0 Then
                        sErr = "Employee Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillA97_Error
                    End If
                    .SetHeadValue "FEmpID", lFItemID
                    .SetHeadValue "FFManagerID", lFItemID
                    .SetHeadValue "FSManagerID", lFItemID
                    
                    '检查Reason Code
                    Set rs = ExecSQL("select FInterID from t_submessage where FTypeID=11400 and fid='MER'", sDsn)
                    If rs.RecordCount <= 0 Then
                        sErr = "Reason Code[" & CStr(varDatas(i)(2)) & "] is NOT created in ERP system."
                        GoTo SynBillA97_Error
                    End If
                    .SetHeadValue "FAdjReason", rs.Fields("FInterID")
                    
                
                
                End If
                
                
                '表体赋值

                .AddRow lRow
                
                ROW = lRow
                
                
                 '检查物料是否存在
'                 lFItemID = GetItemIDByNumber(CNN, CStr(varDatas(i)(5)), 4, lUnitID)
'                 If lFItemID = 0 Then
'                     sErr = "Material Master Data [" & CStr(varDatas(i)(5)) & "] is NOT created in ERP system."
'                     GoTo SynBillA97_Error
'                 End If
'                .SetEntryValue "FItemID", lFItemID, lRow           '物料
'                .SetEntryValue "FUnitID", lUnitID, lRow            '单位
                
                
                rsPOInfo.Filter = "FItemNumber='" & CStr(varDatas(i)(5)) & "'"
                
                '检查物料是否存在
                 If rsPOInfo.RecordCount = 0 Then
                     sErr = "Material Master Data [" & CStr(varDatas(i)(5)) & "] does NOT exist in the original PO."
                     GoTo SynBillA97_Error
                 End If
                .SetEntryValue "FItemID", rsPOInfo.Fields("FItemID").Value, lRow            '物料
                .SetEntryValue "FUnitID", rsPOInfo.Fields("FUnitID").Value, lRow            '单位
                
                '数量必须大于0
                If Val(varDatas(i)(6)) <= 0 Then
                    sErr = "The quantity can NOT be zero."
                    GoTo SynBillA97_Error
                End If
                .SetEntryValue "Fauxqty", Val(varDatas(i)(6)) * lROB, lRow           '数量
                .SetEntryValue "FQty", Val(varDatas(i)(6)) * lROB, lRow          '应收数量
                
                '仓库
                lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 5)
                If lFItemID = 0 Then
                     sErr = "Warehouse Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                     GoTo SynBillA97_Error
                 End If
                .SetEntryValue "FDCStockID", lFItemID, lRow        '收货仓库
                
'                 If Val(rsBillInfo.Fields("FIsStockMgr").Value) <> 0 And Val(rsBillInfo.Fields("FBinID").Value) = 0 Then
'                     sErr = "Bin Master Data [" & rsBillInfo.Fields("FBinCode").Value & "] is NOT created in ERP system or does not in the right bin group."
'                     GoTo SynBillA97_Error
'                 End If
'                .SetEntryValue "FDCSPID", rsBillInfo.Fields("FBinID").Value, lRow        '收货仓位

                 '=========源单信息===========
                 .SetEntryValue "FSourceTranType", 71, lRow          ''源单类型
                 .SetEntryValue "FSourceBillNo", rsPOInfo.Fields("FBillNo").Value, lRow          ''源单编号
                 .SetEntryValue "FSourceInterId", rsPOInfo.Fields("FInterID").Value, lRow          ''源单内码
                 .SetEntryValue "FSourceEntryID", rsPOInfo.Fields("FEntryID").Value, lRow          ''源单行号
'                 .SetEntryValue "FOrderInterID", rsPOInfo.Fields("FInterID").Value, lRow            '
'                 .SetEntryValue "FOrderEntryID", rsPOInfo.Fields("FEntryID").Value, lRow           '
'                 .SetEntryValue "FOrderBillNo", rsPOInfo.Fields("FBillNo").Value, lRow
                '=============================

                .SetEntryValue "FNote", CStr(varDatas(i)(8)), lRow          ''备注
                                
                lRow = lRow + 1
                
NextRow:
'                blnCheck = True ''是否存在有效物料
            Next
            
            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
                checkBillData sDsn, lInterID, lTranType, 0, sErr, 16394 '审核单据
            End If
        End If
    End With

    Set clsBill = Nothing
    Set CNN = Nothing
    SyncStockAdjustment = True
    GetObjectContext.SetComplete

   Exit Function

SynBillA97_Error:
    SyncStockAdjustment = False
    sErr = "Generate ：" & sErr & Err.Description
    GetObjectContext.SetAbort
    Set clsBill = Nothing
    Set CNN = Nothing
    

End Function








