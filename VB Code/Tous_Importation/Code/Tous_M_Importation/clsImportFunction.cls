VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "clsImportFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public rsAccountBinding As ADODB.Recordset '装载帐套绑定配置信息
Public rsItemInfo As ADODB.Recordset '目标帐套的核算项目信息
Public rsAccountInfo As ADODB.Recordset '目标帐套的会计科目信息
Public rsCurrency As ADODB.Recordset '目标帐套的币别信息
Public rsSettle As ADODB.Recordset '目标帐套结算方式
Public rsVoucherGroup As ADODB.Recordset '目标帐套的凭证字
Public rsItemDetail As ADODB.Recordset '源帐套的科目核算项目明细


Private m_conn As ADODB.Connection

'初始化选定的绑定帐套的基础数据数据集
'Private Sub IniDataSet(sdsnB As String)
'    Dim strSQL As String
'
'    Set rsItemInfo = Nothing
'    Set rsAccountInfo = Nothing
'    Set rsCurrency = Nothing
'    Set rsSettle = Nothing
'    Set rsVoucherGroup = Nothing
'    Set rsItemDetail = Nothing
'
'    '取出目标帐套基础资料信息
'    strSQL = "SELECT FItemID,FItemClassID,FNumber FROM T_Item"
'    Set rsItemInfo = ModPub.ExecuteSQL(sdsnB, strSQL)
'    '取出目标帐套科目信息
'    strSQL = "SELECT FAccountID AS FID,FNumber,FDetailID FROM T_Account"
'    Set rsAccountInfo = ModPub.ExecuteSQL(sdsnB, strSQL)
'    '取出目标帐套的币别信息
'    strSQL = "SELECT FCurrencyID AS FID,FNumber FROM t_Currency"
'    Set rsCurrency = ModPub.ExecuteSQL(sdsnB, strSQL)
'    '取出目标帐套的结算方式信息
'    strSQL = "SELECT FItemID AS FID,FNumber FROM t_Settle"
'    Set rsSettle = ModPub.ExecuteSQL(sdsnB, strSQL)
'    '取出目标帐套的凭证字信息
'    strSQL = "SELECT FGroupID AS FID,FName AS FNumber FROM t_VoucherGroup"
'    Set rsVoucherGroup = ModPub.ExecuteSQL(sdsnB, strSQL)
'    '取出源帐套的科目的核算项目明细
'    strSQL = "SELECT T1.FDetailID,T1.FItemClassID,T3.FName AS FClassName,T1.FItemID,T2.FNumber,isnull(t4.FPGU,'')FPGU"
'    strSQL = strSQL & vbCrLf & "FROM t_ItemDetailV T1 LEFT JOIN t_Item T2 ON (T2.FItemID = T1.FItemID AND T2.FItemClassID = T1.FItemClassID)"
'    strSQL = strSQL & vbCrLf & "LEFT JOIN T_ItemClass T3 ON (T3.FItemClassID = T1.FItemClassID)"
'    strSQL = strSQL & vbCrLf & "LEFT JOIN t_Department t4 on t1.FItemID=t4.FItemID WHERE T1.FItemID > 0"
'    Set rsItemDetail = ModPub.ExecuteSQL(sdsnA, strSQL)
'
'    '取出目标帐套User信息
'    strSQL = "SELECT FUserID AS FID,FName as FNumber FROM t_user"
'    Set rsUser = ModPub.ExecuteSQL(sdsnB, strSQL)
'End Sub

Public Function CreateVouchers(dsn As String, strUUID As String, ByRef msg As StringBuilder) As Boolean
    Dim objEBSGLV As Object
    Dim result As Integer
    Dim i As Long
    
    Dim rsVoucher As ADODB.Recordset
    Dim rsVoucherData As ADODB.Recordset
    Dim rsError As ADODB.Recordset
    Dim strSQL As String
    
    Dim objSave As Object                           '创建凭证接口
    Dim dicVoucher As KFO.Dictionary                '凭证单条记录
    Dim vctVoucher As KFO.Vector                    '凭证记录集合
    Dim dicVoucherEntry As KFO.Dictionary           '凭证体
    Dim vctVouncherEntry As KFO.Vector              '凭证体集合
    Dim vctDetail As KFO.Vector                     '核算项目
    Dim dicDetail As KFO.Dictionary                 '核算项目
    Dim vctResult As KFO.Vector                     '执行结果
    Dim strMsgVoucherCreate As String               '记录错误信息
    Dim VoucherIDB As Long                          '记录目标帐套新生成的凭证内码
    
    Dim amount
    Dim amountFor
    Dim dblTotalDebitAmount As Double
    Dim dblTotalCreditAmount As Double
    
    GetObjectContext.SetAbort
On Error GoTo Err
    Dim p As New CParse
    p.ParseString dsn
    
    EnsureDBConnection dsn
    
    Set objSave = CreateObject("BillDataAccess.GetData")
    
    Dim rs As ADODB.Recordset '辅助记录集
    
    '查询本位币的金额精度
    Dim cyBaseDecimal As Long
    Set rs = ExecSQL("Select FScale From t_Currency Where FCurrencyID=1")
    cyBaseDecimal = rs.Fields("FScale").Value
    Set rs = Nothing
    
    '取凭证头
    strSQL = "select FVoucherCategory,FVchNumber,FYear,FPeriod from PF_t_VoucherData where FUUID='" & strUUID & "' "
    strSQL = strSQL & "group by FVoucherCategory,FVchNumber,FYear,FPeriod"
            
    '在源帐套中取出所有未传送的（或传送没有成功的）凭证内码
    Set rsVoucher = objSave.ExecuteSQL(dsn, strSQL)
    If rsVoucher.RecordCount = 0 Then
        Set rsVoucher = Nothing
        msg.Append "There is no voucher data"
        Exit Function
    End If
    
'    IniDataSet
    
    '取凭证所有数据
    strSQL = "select * from PF_t_VoucherData where FUUID='" & strUUID & "'"
    Set rsVoucherData = objSave.ExecuteSQL(dsn, strSQL)
    
    rsVoucher.MoveFirst
    Set vctResult = Nothing
    Set vctVoucher = Nothing
    Set vctVoucher = New KFO.Vector
    While Not rsVoucher.EOF
        rsVoucherData.Filter = "FVchNumber = '" & CStr(rsVoucher.Fields("FVchNumber").Value) & "' AND FYear = '" & CStr(rsVoucher.Fields("FYear").Value) & _
        "' AND FPeriod = '" & CStr(rsVoucher.Fields("FPeriod").Value) & "' AND FVoucherCategory = '" & CStr(rsVoucher.Fields("FVoucherCategory").Value) & "'"
        rsVoucherData.MoveFirst
        
        '====================Voucher head 凭证头=====================
        Set dicVoucher = New KFO.Dictionary
        Set vctVouncherEntry = New KFO.Vector
        dicVoucher("FNumber") = rsVoucherData.Fields("FVchNumber").Value                                         '凭证号
        dicVoucher("FExplanation") = rsVoucherData.Fields("FReference").Value                                    '摘要
        dicVoucher("FHandler") = 16394                                                                           '经办
        dicVoucher("FTransDate") = rsVoucherData.Fields("FTransactionDate").Value                                '业务日期

'        lTempID = GetIDFromB(rsUser, rsVoucherData.Fields("FApproveName").Value)
'        dicVoucher("FApproveID") = IIf(lTempID = 0, 16394, lTempID)                                             '审批
        
'        dicVoucher("FFootNote") = rsVHead("FFootNote").Value                                                '批注
'        dicVoucher("FTranType") = rsVoucherData.Fields("FTranType").Value                                       '单据类型
        dicVoucher("FDate") = rsVoucherData.Fields("FDate").Value                                                '凭证日期
'        dicVoucher("FInternalInd") = rsVoucherData.Fields("FinternalInd").Value                                 '凭证机制
        dicVoucher("FGroupID") = rsVoucherData.Fields("FVoucherCategory").Value                                  '凭证字
'        dicVoucher("FVoucherGroupName") = rsVoucherData.Fields("FVoucherCategory").Value
'        dicVoucher("FYear") = rsVoucherData.Fields("FYear").Value
'        dicVoucher("FPeriod") = rsVoucherData.Fields("FPeriod").Value
        
        dicVoucher("FPreparerID") = p.UserID                                                                     '制单人
        dicVoucher("FReference") = rsVoucherData.Fields("FReference").Value                                      '参考
        dicVoucher("FEntryCount") = rsVoucherData.RecordCount                           '凭证头分录数
        
'        FVoucherID = rsVoucherData.Fields("FVoucherID").Value
'        FYear = rsVoucherData.Fields("FYear").Value
'        FPeriod = rsVoucherData.Fields("FPeriod").Value
'        FGroupID = rsVoucherData.Fields("FGroupID").Value
        Dim rsAcc As ADODB.Recordset
        Dim FAccountID As Long
        While Not rsVoucherData.EOF
        
        '====================Voucher entry 凭证体=====================
            Set dicVoucherEntry = New KFO.Dictionary
            FAccountID = Val(rsVoucherData.Fields("FAccountNumber").Value)
            dicVoucherEntry("FAccountID") = rsVoucherData.Fields("FAccountNumber").Value                        '科目
            Set rsAcc = ExecSQL("Select (Select FCurrencyID From t_Account where FAccountID=" & FAccountID & ") FCurrencyID," _
                & "(Select isnull(FExchangeRate,0) From t_Currency where FCurrencyID =" & rsVoucherData.Fields("FCurrency").Value & " ) FExchangeRate " _
                & ",(Select isnull(FScale,0) From t_Currency where FCurrencyID =" & rsVoucherData.Fields("FCurrency").Value & " ) FScale ")
            dicVoucherEntry("FExplanation") = rsVoucherData.Fields("FDescription").Value                       '摘要
            '如果是多币种核算或支持指定的币种,就计算本位币值
            If rsAcc.Fields("FCurrencyID").Value = 0 Or rsAcc.Fields("FCurrencyID").Value = Val(rsVoucherData.Fields("FCurrency").Value) Then
                dicVoucherEntry("FCurrencyID") = rsVoucherData.Fields("FCurrency").Value                           '币别
                dicVoucherEntry("FExchangeRate") = Val(rsAcc.Fields("FExchangeRate").Value)
            ElseIf rsAcc.Fields("FCurrencyID").Value = 1 Then '如果此K3科目只支持本位币，则需要将原转成本位币
                dicVoucherEntry("FCurrencyID") = 1
                dicVoucherEntry("FExchangeRate") = 1
            Else
                Err.Raise 1, "CreateVouchers", "Account does not support currency that from data!"
            End If

            
            If Val(rsVoucherData.Fields("FDebitAmount").Value) <> 0 Then '注意防范红字凭证
                dicVoucherEntry("FDC") = 1
                amountFor = KDRound(rsVoucherData.Fields("FDebitAmount").Value, rsAcc.Fields("FScale").Value)
                If rsVoucherData.Fields("FCurrency") = 1 Then
                    amount = amountFor
                Else
                    amount = KDRound(amountFor * rsAcc.Fields("FExchangeRate").Value, cyBaseDecimal)
                    '如果科目是本币核算的，但导入数据是原币的,则换算成本币
                    If dicVoucherEntry("FCurrencyID") = 1 Then
                        amountFor = amount
                    End If
                End If
                dicVoucherEntry("FAmount") = amount
                dicVoucherEntry("FAmountFor") = amountFor
                dblTotalDebitAmount = dblTotalDebitAmount + amount
            Else
                dicVoucherEntry("FDC") = 0
                amountFor = KDRound(rsVoucherData.Fields("FCreditAmount").Value, rsAcc.Fields("FScale").Value)
                If rsVoucherData.Fields("FCurrency") = 1 Then
                    amount = KDRound(amountFor, cyBaseDecimal)
                Else
                    amount = KDRound(amountFor * rsAcc.Fields("FExchangeRate").Value, cyBaseDecimal)
                    If dicVoucherEntry("FCurrencyID") = 1 Then
                        amountFor = amount
                    End If
                End If
                dicVoucherEntry("FAmount") = amount
                dicVoucherEntry("FAmountFor") = amountFor
                dblTotalCreditAmount = dblTotalCreditAmount + amount
            End If
'            dicVoucherEntry("FEntryID") = rsVoucherData.Fields("FEntryID").Value                               '凭证分录号
'            dicVoucherEntry("FDetailID") = 0                                                                   '核算项目
            dicVoucherEntry("Fquantity") = 0                                                                    '数量
'            If rsVoucherData.Fields("FMeasureUnitID").Value = 0 Then
                dicVoucherEntry("FMeasureUnitID") = 0
'            Else
'                dicVoucherEntry("FMeasureUnitID") = GetItemIDFromB(rsItemInfo, 7, rsVoucherData.Fields("FUnitNumber").Value)         '数量单位
'            End If
            dicVoucherEntry("FUnitPrice") = 0                                                                   '单价
'            dicVoucherEntry("FInternalInd") = rsVoucherData.Fields("FInternalInd").Value                        '机制凭证
            dicVoucherEntry("FSettleTypeID") = 0                                                                 '结算方式
'            dicVoucherEntry("FSettleNo") = rsVoucherData.Fields("FSettleNo").Value                              '结算号
'            dicVoucherEntry("FTransNo") = rsVoucherData.Fields("FTransNo").Value                                '往来业务编号
        '============================================科目1========================================

            Set vctDetail = New KFO.Vector
            '如果有客户核算项目则添加
            If Len(rsVoucherData.Fields("FAICustomer").Value) > 0 Then
                Set dicDetail = New KFO.Dictionary
                dicDetail("FItemClassID") = 1
                dicDetail("FItemID") = rsVoucherData.Fields("FAICustomer").Value
                vctDetail.Add dicDetail
            End If
            '如果有物料核算项目则添加
            If Len(rsVoucherData.Fields("FAIMaterial").Value) > 0 Then
                Set dicDetail = New KFO.Dictionary
                dicDetail("FItemClassID") = 4
                dicDetail("FItemID") = rsVoucherData.Fields("FAIMaterial").Value
                vctDetail.Add dicDetail
            End If
            If Not vctDetail Is Nothing Then
                If vctDetail.Size > 0 Then
                    Set dicVoucherEntry("_Details") = vctDetail
                End If
            End If
            '将单条记录封装到集合体中
            vctVouncherEntry.Add dicVoucherEntry
            
            Set dicVoucher("_Entries") = vctVouncherEntry                    '保存凭证体
            
            rsVoucherData.MoveNext
        Wend ' --rsVoucherData
        
        dicVoucher("FDebitTotal") = dblTotalDebitAmount          '凭证头总借方金额
        dicVoucher("FCreditTotal") = dblTotalCreditAmount         '凭证头总贷方金额
        
        dblTotalDebitAmount = 0
        dblTotalCreditAmount = 0
        
        vctVoucher.Add dicVoucher
        
        rsVoucher.MoveNext
        
    Wend '--rsVoucher
    

    Set objEBSGLV = CreateObject("EBSGLVoucher.VoucherUpdate")
    For i = 1 To vctVoucher.Size
        Set dicVoucher = vctVoucher(i)
        result = objEBSGLV.Create(dsn, dicVoucher)
        vctVoucher(i)("VoucherID") = result
    Next i
    
    CreateVouchers = True
    Set objEBSGLV = Nothing
    Set dicVoucher = Nothing
    
    GetObjectContext.SetComplete
    Exit Function
Err:
    GetObjectContext.SetAbort
    msg.Append Err.Source & ":" & Err.Description
'    strSql = "delete from PF_t_VoucherData where FUUID = '" & strUUID & "'"
'    Set rsError = objSave.ExecuteSQL(dsn, strSql)
    Err.Clear
    Set objEBSGLV = Nothing
    CreateVouchers = False
    Set dicVoucher = Nothing
    
End Function


Public Function CreateSTN(dsn As String, strUUID As String, ByRef strBillNo As String, ByRef msg As String) As Boolean
    Dim oDataSrv As Object
    Dim oTableInfo As Object
    Dim oData As Object
    
    Dim oSrv As Object
    Dim objSave As Object
    
    Dim strSQL As String
    Dim rsHead As Recordset
    Dim rsEntry As Recordset
    Dim i As Integer
    Dim j As Integer
    
    Dim strTempMsg As String
    Dim process As String
    
    Dim objResult As Object
    
On Error GoTo Err
    GetObjectContext.SetAbort
    
    process = "Preparing database conneciton..."
    Call EnsureDBConnection(dsn)
    
    process = "Creating  BillDataAccess.GetData object instance..."
    Set objSave = CreateObject("BillDataAccess.GetData")


    process = "Creating K3ClassTpl.DataSrv instance..."
    Set oDataSrv = CreateObject("K3ClassTpl.DataSrv")
    oDataSrv.PropString = dsn
    oDataSrv.ClassTypeID = 200000101
    Set oTableInfo = oDataSrv.TableInfo

    '取是本次事务中在系统中尚未导入过的不重复的发票号码
    process = "Get header from tempary table..."
    strSQL = "select distinct FStockOutID,FStockInID from t_Tous_STNImportData where FUUID='" & strUUID & "'"
    Set rsHead = objSave.ExecuteSQL(dsn, strSQL)
    
    If rsHead.RecordCount <= 0 Then '没有需要导入的数据
        msg = "No stock transfer notice has been imported."
        CreateSTN = True
        Exit Function
    End If
    
    '取本次事务的发票数据
    process = "Get detail from tempary table..."
    strSQL = "select * from t_Tous_STNImportData where FUUID='" & strUUID & "'"
    Set rsEntry = objSave.ExecuteSQL(dsn, strSQL)
    
    process = "Package data..."
    Dim vecSTN As New KFO.Vector
    For i = 1 To rsHead.RecordCount
        rsEntry.Filter = "FStockOutID='" & rsHead.Fields("FStockOutID").Value & "' and FStockInID='" & rsHead.Fields("FStockInID").Value & "'"
'        oDataSrv.PropString = dsn
'        oDataSrv.ClassTypeID = 200000101
        Set oData = oDataSrv.GetEmptyBill()
        PackageSTNData dsn, oData, oTableInfo, rsEntry
        vecSTN.Add oData
        rsHead.MoveNext
    Next
    
    process = "Saving STN to k3 using k3 insterface..."
    Set oSrv = CreateObject("K3MClass.BillUpdate")
    For i = 1 To vecSTN.Size
        Set oData = vecSTN(i)
'        objResult = oSrv.Save(dsn, oDataSrv.classType, oDataSrv.classTypeEntry, oTableInfo, oData, oDataSrv.dctLink)
        oSrv.Save dsn, oDataSrv.classType, oDataSrv.classTypeEntry, oTableInfo, oData, oDataSrv.dctLink
        strBillNo = strBillNo & oData("Page1")(oTableInfo("map")("FBillNo"))("FFLD") & ","
'        Set objResult = oData
    Next
    
    m_conn.Close
    Set m_conn = Nothing
    GetObjectContext.SetComplete
    CreateSTN = True
    Exit Function
Err:
    msg = process & ":" & Err.Source & ":" & Err.Description
    GetObjectContext.SetAbort
    Err.Clear
    CreateSTN = False
End Function

Private Sub PackageSTNData(dsn As String, ByRef oData As Object, oTableInfo As Object, rsEntry As ADODB.Recordset)
    oData("Page1")(oTableInfo("map")("FClassTypeID"))("FFLD") = 200000101
    oData("Page1")(oTableInfo("map")("FDate"))("FFLD") = Format(Now, "yyyy-mm-dd")
'    oData("Page1")(oTableInfo("map")("FBillNo"))("FFLD") = rsEntry.Fields("FInvNumber").Value
    
    Dim p As New CParse
    p.ParseString dsn
    oData("Page1")(oTableInfo("map")("FBiller"))("FFLD") = p.UserID
    oData("Page1")(oTableInfo("map")("FSCStockID"))("FFLD") = rsEntry.Fields("FStockOutID").Value
    oData("Page1")(oTableInfo("map")("FDCStockID"))("FFLD") = rsEntry.Fields("FStockInID").Value
    
    '写明细表
    Dim oPage2 As Object
    Dim j As Long
    For j = 1 To rsEntry.RecordCount - 1
        Set oPage2 = oData("Page2")(j).CloneAll
        oData("Page2").Add oPage2
    Next
    
    For j = 1 To rsEntry.RecordCount
        oData("Page2")(j)(oTableInfo("map")("FEntryID2"))("FFLD") = j
        
        oData("Page2")(j)(oTableInfo("map")("FItemID"))("FFLD") = rsEntry.Fields("FItemID").Value
        oData("Page2")(j)(oTableInfo("map")("FTransferQty"))("FFLD") = rsEntry.Fields("FQty").Value
        
        rsEntry.MoveNext
    Next
    
        
End Sub

Public Function CreateMaterial(dsn As String, strUUID As String, ByRef msg As String) As Boolean

    Dim objSave As Object                           '创建凭证接口
    Dim oSave As Object
    Dim rs As Recordset
    Dim i As Integer
    Dim item As KFO.Dictionary
    Dim lProductID As Long
    Dim lAccountID  As Long
    Dim strSQL As String
        
On Error GoTo Err
    GetObjectContext.SetAbort
    Set objSave = CreateObject("BillDataAccess.GetData")
    
    '取9999的AccountID
    strSQL = "select FAccountID from t_Account where FNumber ='9999'"
    Set rs = objSave.ExecuteSQL(dsn, strSQL)
    If rs.RecordCount = 0 Then
        msg = msg & "Not found accounting subjects 9999 in K3" & vbCrLf
        CreateMaterial = False
        Set objSave = Nothing
        Set rs = Nothing
        Exit Function
    Else
        lAccountID = rs.Fields("FAccountID").Value
    End If
    
    Set rs = Nothing

    '取所有不存在的物料，然后自动生成
    strSQL = "select t1.FProductNumber,Max(isnull(t1.FProductName,'')) FProductName,Max(isnull(t1.FUoM,0)) FUoM,Max(isnull(t3.FUnitGroupID,0)) FUnitGroupID from PF_t_InvoiceData t1"
    strSQL = strSQL & vbCrLf & "left join t_ICItem t2 on t1.FProductNumber=t2.FNumber"
    strSQL = strSQL & vbCrLf & "left join t_MeasureUnit t3 on t1.FUoM=t3.FItemID"
    strSQL = strSQL & vbCrLf & "where t2.FItemID is null and t1.FUUID='" & strUUID & "' "
    strSQL = strSQL & vbCrLf & "group by t1.FProductNumber"
    Set rs = objSave.ExecuteSQL(dsn, strSQL)
    
    If rs.RecordCount <= 0 Then
        CreateMaterial = True
        Exit Function
    End If
    
    For i = 1 To rs.RecordCount
        Set item = New KFO.Dictionary
        
        item("FNumber") = rs.Fields("FProductNumber").Value
        item("FName") = rs.Fields("FProductName").Value
        item("FFullName") = rs.Fields("FProductName").Value
        
        '物料属性
        item("FErpClsID") = 1
        '计价方法
        item("FTrack") = 76
        
        '单位
        item("FUnitGroupID") = rs.Fields("FUnitGroupID").Value
        item("FUnitID") = rs.Fields("FUoM").Value
        item("FOrderUnitID") = rs.Fields("FUoM").Value
        item("FProductUnitID") = rs.Fields("FUoM").Value
        item("FSaleUnitID") = rs.Fields("FUoM").Value
        item("FStoreUnitID") = rs.Fields("FUoM").Value
        item("FSecUnitID") = 0
        
        '科目
        item("FAcctID") = lAccountID
        item("FSaleAcctID") = lAccountID
        item("FCostAcctID") = lAccountID
                
        '控制类型
        item("FCtrlType") = 14039
        
        '检验方式
        item("FInspectionLevel") = 352
        item("FProChkMde") = 352
        item("FSOChkMde") = 352
        item("FStkChkMde") = 352
        item("FOtherChkMde") = 352
        item("FWthDrwChkMde") = 352
        item("FWWChkMde") = 352
        
        item("FDetail") = True
        
        
        item("FPriceDecimal") = 5
        item("FQtyDecimal") = 2
        Set oSave = CreateObject("EBSGLItem.ItemUpdate")
        
        lProductID = oSave.Create(dsn, 4, item)
        Set item = Nothing
        If lProductID <= 0 Then GoTo Err
        
        rs.MoveNext
    Next
    

    CreateMaterial = True
    GetObjectContext.SetComplete
    Exit Function
Err:
    CreateMaterial = False
    msg = Err.Description
    GetObjectContext.SetAbort
    Err.Clear
    
End Function

Public Function CreateCustomer(dsn As String, strUUID As String, ByRef msg As String) As Boolean
    
    Dim objSave As Object                           '创建凭证接口
    Dim oSave As Object
    Dim rs As Recordset
    Dim i As Integer
    Dim item As KFO.Dictionary
    Dim lCustomerID As Long
    Dim strSQL As String
    
On Error GoTo Err

    GetObjectContext.SetAbort
    Set objSave = CreateObject("BillDataAccess.GetData")
    
    '取所有不存在的客户，然后自动生成
    strSQL = "select t1.FCustomerNumber,Max(t1.FCustomerName) FCustomerName,Max(isnull(t1.FTaxRegistration,'')) FTaxRegistration,"
    strSQL = strSQL & vbCrLf & "Max(isnull(t1.FContactPerson,'')) FContactPerson,Max(isnull(t1.FTelNumber,'')) FTelNumber,"
    strSQL = strSQL & vbCrLf & "Max(isnull(t1.FAddress1,'')) FAddress1 , Max(isnull(t1.FAddress2,'')) FAddress2, Max(isnull(t1.FAddress3,'')) FAddress3, Max(isnull(t1.FAddress4,'')) FAddress4,"
    strSQL = strSQL & vbCrLf & "Max(isnull(t1.FPostcode,'')) FPostCode, Max(isnull(t1.FMailAddress,'')) FMailAddress, Max(isnull( t1.FBank,'')) FBank,Max(isnull( t1.FBankAccount,'')) FBankAccount"
    strSQL = strSQL & vbCrLf & "from PF_t_InvoiceData t1 left join t_Organization t2 on t1.FCustomerNumber=t2.FNumber"
    strSQL = strSQL & vbCrLf & "where t2.FItemID is null and t1.FUUID='" & strUUID & "' "
    strSQL = strSQL & vbCrLf & "group by t1.FCustomerNumber"
    Set rs = objSave.ExecuteSQL(dsn, strSQL)
    
    If rs.RecordCount <= 0 Then
        CreateCustomer = True
        Exit Function
    End If
    
    For i = 1 To rs.RecordCount
        Set item = New KFO.Dictionary
        
        item("FNumber") = rs.Fields("FCustomerNumber").Value
        item("FName") = rs.Fields("FCustomerName").Value
        item("FFullName") = rs.Fields("FCustomerName").Value
        
        '增值税率
        item("FValueAddRate") = 17
        
        item("FTaxNum") = rs.Fields("FTaxRegistration").Value
        item("FContact") = rs.Fields("FContactPerson").Value
        item("FPhone") = rs.Fields("FTelNumber").Value
        item("FAddress1") = rs.Fields("FAddress1").Value
        item("FAddress2") = rs.Fields("FAddress2").Value
        item("FAddress3") = rs.Fields("FAddress3").Value
        item("FAddress4") = rs.Fields("FAddress4").Value
        item("FPostalCode") = rs.Fields("FPostcode").Value
        item("FEmail") = rs.Fields("FMailAddress").Value
        item("FBank") = rs.Fields("FBank").Value
        item("FAccount") = rs.Fields("FBankAccount").Value
        
        item("FDetail") = True
        
        Set oSave = CreateObject("EBSGLItem.ItemUpdate")
        
        lCustomerID = oSave.Create(dsn, 1, item)
        Set item = Nothing
        If lCustomerID <= 0 Then GoTo Err
        
        rs.MoveNext
    Next
    
    CreateCustomer = True
    
    GetObjectContext.SetComplete
    Exit Function
Err:
    CreateCustomer = False
    msg = Err.Description & msg
    GetObjectContext.SetAbort
    Err.Clear
    
End Function

'代码转ID
'特殊情况:币别除了转换成ID之外，还需要得到汇率。因此参数dblRate用来返回汇率
Private Function GetItemID(dsn As String, strNumber As String, lngItemClassID As Long, Optional ByRef dblRate As Double) As Long
    Dim objSave As Object
    Dim rs As Recordset
    Dim strSQL As String
On Error GoTo Err
    Set objSave = CreateObject("BillDataAccess.GetData")
    
    If lngItemClassID = 1 Then
        '客户
        strSQL = "select FItemID from t_Organization where FNumber='" & strNumber & "'"
        Set rs = objSave.ExecuteSQL(dsn, strSQL)
        
        If rs.RecordCount > 0 Then
            GetItemID = CLng(rs.Fields("FItemID").Value)
        Else
            GetItemID = 0
        End If
        
    ElseIf lngItemClassID = 4 Then
        '物料
        strSQL = "select FItemID from t_icitem where FNumber='" & strNumber & "'"
        Set rs = objSave.ExecuteSQL(dsn, strSQL)
        
        If rs.RecordCount > 0 Then
            GetItemID = CLng(rs.Fields("FItemID").Value)
        Else
            GetItemID = 0
        End If
    ElseIf lngItemClassID = 99 Then
        '币别
        strSQL = "select FCurrencyID,FExchangeRate from t_Currency where FNumber='" & strNumber & "'"
        Set rs = objSave.ExecuteSQL(dsn, strSQL)
        
        If rs.RecordCount > 0 Then
            dblRate = CDbl(rs.Fields("FExchangeRate").Value)
            GetItemID = CLng(rs.Fields("FCurrencyID").Value)
        Else
            dblRate = 0
            GetItemID = 0
        End If
    End If

    Exit Function
Err:
    dblRate = 0
    GetItemID = 0
End Function


Private Sub EnsureDBConnection(dsn As String)
    
    If Not m_conn Is Nothing Then Exit Sub
    
    Dim db As New CDataSource
    Dim parse As New CParse
    parse.ParseString dsn
    Set db.ParseObject = parse
    Dim rs As ADODB.Recordset
    Set m_conn = db.Connection
    m_conn.CommandTimeout = 0
End Sub

Private Function ExecSQL(sqlCommandText As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open sqlCommandText, m_conn, adOpenStatic, adLockBatchOptimistic
    Set ExecSQL = rs
End Function
