VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "HQ_05"
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
'    strLanguage = datasource.ParseObject.LanguageID
End Function

'---------------------------------------------------------------------------------------
' Procedure : ImportDeliveries
' DateTime  :
' Author    :
' Purpose   : 生成采购订单 (HQ TO ERP)
'---------------------------------------------------------------------------------------
Public Function ImportDeliveries(ByVal sDsn As String, ByVal dicHeader As KFO.Dictionary, ByVal vec As KFO.Vector, ByRef sErr As String) As Boolean
    Dim CNN As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim rsPOInfo As ADODB.Recordset
    Dim clsBill As AutoOldBill
    Dim bRet As Boolean
    Dim lRow As Long
    Dim i As Long, lInterID As Long

    Dim lFItemID As Long
    Dim lUnitID As Long
    Dim dRate As Double
    
    Dim dblFExchangeRate As Double, lFCurrencyID As Long

    Dim lROB As Long, ROW As Long
    
    Dim strSQL As String, sBillNo As String
    Dim dic As KFO.Dictionary
    
On Error GoTo Err

    InitDataEnv sDsn
    Set CNN = datasource.Connection

    Set clsBill = New AutoOldBill

    lRow = 1
    lROB = 1
    GetObjectContext.SetAbort
    
    With clsBill
        .sDsn = sDsn
        bRet = .NewBill(71)

        If bRet = True Then
                
            '表头赋值
            .SetHeadValue "FPOMode", 36680 '采购模式:  默认为普通采购
            .SetHeadValue "FPOStyle", 252  '采购方式: 默认为赊购
            .SetHeadValue "FPlanCategory", 1
                
            '检查POS Number是否存在，存在则不允许重复导入
            Set rs = ExecSQL("select 1 from POOrder where isnull(FHQNumber,'')='" & CStr(dicHeader("DocNumber")) & "'", sDsn)
            If rs.RecordCount > 0 Then
                sErr = "Document Number[" & CStr(dicHeader("DocNumber")) & "] has already existed in ERP system."
                GoTo Err
            End If
                        
            .SetHeadValue "FHQNumber", CStr(dicHeader("DocNumber"))   '总部编号
            .SetHeadValue "FDate", dicHeader("Date") 'CStr(varDatas(i)(4))  '日期
            
            '采购类型
            If CStr(dicHeader("DocNumber")) <> "50" Then
                .SetHeadValue "FPurchaseType", 11101
            Else
                .SetHeadValue "FPurchaseType", 11102
            End If
            
            
            '币别汇率，默认欧元
            dblFExchangeRate = GetExchangeRate(CNN, "EUR", dicHeader("Date"), lFCurrencyID)
            If lFCurrencyID = 0 Then
                sErr = "Currency [EUR] is not found."
                GoTo Err
            End If
            If dblFExchangeRate = 0 Then
                sErr = "No available exchange rate for Currency [EUR]."
                GoTo Err
            End If
            .SetHeadValue "FExchangeRateType", 1
            .SetHeadValue "FCurrencyID", lFCurrencyID
            .SetHeadValue "FExchangeRate", dblFExchangeRate
                
            
            '检查供应商是否存在
            '供应商默认为Tous HQ，代码为50
            lFItemID = GetItemID(CNN, "50", 8)
            If lFItemID = 0 Then
                sErr = "Supplier Master Data [50] is NOT created in ERP system."
                GoTo Err
            End If
            .SetHeadValue "FSupplyID", lFItemID
            .SetHeadValue "FHQOrigin", CStr(dicHeader("Supplier"))
            
            '检查部门是否存在
            lFItemID = GetItemID(CNN, CStr(dicHeader("Dept")), 2)
            If lFItemID = 0 Then
                sErr = "Department Master Data [" & CStr(dicHeader("Dept")) & "] is NOT created in ERP system."
                GoTo Err
            End If
            .SetHeadValue "FDeptID", lFItemID
            
            
            '检查职员是否存在
            lFItemID = GetItemID(CNN, CStr(dicHeader("Dept")), 3)
            If lFItemID = 0 Then
                sErr = "Employee Master Data [" & CStr(dicHeader("Dept")) & "] is NOT created in ERP system."
                GoTo Err
            End If
            .SetHeadValue "FEmpID", lFItemID
            .SetHeadValue "FFManagerID", lFItemID
            .SetHeadValue "FSManagerID", lFItemID
            
            '检查币别
            Set rs = ExecSQL("select FCurrencyID from t_Currency where FNumber='" & CStr(dicHeader("Currency")) & "'", sDsn)
            If rs.RecordCount <= 0 Then
                sErr = "Currency Code[" & CStr(dicHeader("Currency")) & "] is NOT created in ERP system."
                GoTo Err
            End If
            .SetHeadValue "FCurrencyID", rs.Fields("FCurrencyID")
                
                    
            '表体赋值
            For i = 1 To vec.Size
                Set dic = vec(i)
                
                .AddRow lRow
                
                ROW = lRow
                
                 '检查物料是否存在
                 lFItemID = GetItemIDByNumber(CNN, CStr(dic("SKU")), 4, lUnitID)
                 If lFItemID = 0 Then
                     sErr = "Material Master Data [" & CStr(dic("SKU")) & "] is NOT created in ERP system."
                     GoTo Err
                 End If
                .SetEntryValue "FItemID", lFItemID, lRow           '物料
                .SetEntryValue "FUnitID", lUnitID, lRow            '单位
                
                '数量必须大于0
                If Val(dic("Qty")) <= 0 Then
                    sErr = "The quantity should be positive number."
                    GoTo Err
                End If
                .SetEntryValue "FAuxQty", Val(dic("Qty")) * lROB, lRow          '数量
                .SetEntryValue "FQty", Val(dic("Qty")) * lROB, lRow          '数量
                
                .SetEntryValue "FCess", 0, lRow          '税率
                .SetEntryValue "FTaxAmount", 0, lRow          '税额
                
                .SetEntryValue "Fauxprice", Val(dic("Price")) * lROB, lRow          '单价
                .SetEntryValue "FAuxTaxPrice", Val(dic("Price")) * lROB, lRow          '单价
                .SetEntryValue "FAuxPriceDiscount", Val(dic("Price")) * lROB, lRow          '单价
                
                .SetEntryValue "FAmount", Val(dic("Amount")) * lROB, lRow          '金额
                .SetEntryValue "FAllAmount", Val(dic("Amount")) * lROB, lRow       '价税合计
                
                .SetEntryValue "FDate", dicHeader("Date"), lRow  '交货日期
                                                
                lRow = lRow + 1
                
            Next
            
            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
                checkBillData sDsn, lInterID, 71, 0, sErr, 16394 '审核单据
                
            End If
        End If
    End With

    Set clsBill = Nothing
    Set CNN = Nothing
    ImportDeliveries = True
    GetObjectContext.SetComplete

   Exit Function

Err:
    ImportDeliveries = False
    sErr = sErr & Err.Description
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







