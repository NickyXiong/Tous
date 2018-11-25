VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************
'Function: 生成调拨单并审核
'Date: 2013/11/08
'*******************************************************

Option Explicit

Private datasource As CDataSource
Private m_oParse As CParse
Private strLanguage As String


''/* 销售出库单 Sales
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
' Procedure : SynBillD01
' DateTime  : 2013-1-25 00:05
' Author    :
' Purpose   : 生成调拨单并审核
'---------------------------------------------------------------------------------------
'
Public Function SynBillD01(ByVal sDsn As String, varDatas As Variant, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String) As Boolean
    Dim CNN As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim clsBill As AutoOldBill
    Dim bRet As Boolean
    Dim lRow As Long
    Dim i As Long

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
    
    Dim lInTransitWH As Long

On Error GoTo SynBillD01_Error

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
    

    lFItemID = GetItemID(CNN, "T-WH", 5)
    If lFItemID = 0 Then
         sErr = "Warehouse Master Data [T-WH] is NOT created in ERP system."
         GoTo SynBillD01_Error
    End If
    lInTransitWH = lFItemID
    
    With clsBill
        .sDsn = sDsn

        bRet = .NewBill(41)

        If bRet = True Then
            For i = LBound(varDatas) To UBound(varDatas)

                If i = 0 Then '表头
                
                    If Val(varDatas(i)(3)) = -1 Then
                        '检查POS Number是否存在，存在则不允许重复导入
                        '当类型是调出时，检查调入仓库为T-WH的是否存在，如果存在则不允许导入
                        strSQL = "select t3.FItemID from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID "
                        strSQL = strSQL & vbCrLf & "inner join t_Stock t3 on t2.FDCStockID=t3.FItemID"
                        strSQL = strSQL & vbCrLf & "where t1.FTranType=41 and isnull(t1.FPOSNumber,'')='" & CStr(varDatas(i)(1)) & "' and t3.FNumber='T-WH'"
                        Set rs = ExecSQL(strSQL, sDsn)
                        If rs.RecordCount > 0 Then
                            sErr = "Transfer-out: POS Number[" & CStr(varDatas(i)(1)) & "] has existed in ERP system."
                            GoTo SynBillD01_Error
'                        Else
'                            lInTransitWH = rs.Fields("FItemID").Value
                        End If
                        Set rs = Nothing
                               
                    ElseIf Val(varDatas(i)(3)) = 1 Then
                        '检查POS Number是否存在，存在则不允许重复导入
                        '当类型是调入时，检查调出仓库为T-WH的是否存在，如果存在则不允许导入
                        strSQL = "select t3.FItemID from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID "
                        strSQL = strSQL & vbCrLf & "inner join t_Stock t3 on t2.FSCStockID=t3.FItemID"
                        strSQL = strSQL & vbCrLf & "where t1.FTranType=41 and isnull(t1.FPOSNumber,'')='" & CStr(varDatas(i)(1)) & "' and t3.FNumber='T-WH'"
                        Set rs = ExecSQL(strSQL, sDsn)
                        If rs.RecordCount > 0 Then
                            sErr = "Transfer-in: POS Number[" & CStr(varDatas(i)(1)) & "] has existed in ERP system."
                            GoTo SynBillD01_Error
'                        Else
'                            lInTransitWH = rs.Fields("FItemID").Value
                        End If
                        Set rs = Nothing
                        
                        '当类型是调入时，检查调入仓库为T-WH的是否存在，如果不存在则不允许导入（代表入库先于出库）
                        strSQL = "select 1 from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID "
                        strSQL = strSQL & vbCrLf & "inner join t_Stock t3 on t2.FDCStockID=t3.FItemID"
                        strSQL = strSQL & vbCrLf & "where t1.FTranType=41 and isnull(t1.FPOSNumber,'')='" & CStr(varDatas(i)(1)) & "' and t3.FNumber='T-WH'"
                        Set rs = ExecSQL(strSQL, sDsn)
                        If rs.RecordCount <= 0 Then
                            sErr = "Transfer-in: POS Number[" & CStr(varDatas(i)(1)) & "] can NOT find transfer-out record."
                            GoTo SynBillD01_Error
                        End If
                        Set rs = Nothing
                        
                    End If
                
     
                    .SetHeadValue "FDate", Format(CDate(varDatas(i)(4)), "yyyy-mm-dd")  '日期
                    .SetHeadValue "FPOSNumber", CStr(varDatas(i)(1)) 'Format(CDate(varDatas(i)(4)), "yyyy-mm-dd")  '日期
    
                    .SetHeadValue "FSourStore", CStr(varDatas(i)(9))   '调出门店
                    .SetHeadValue "FDestStore", CStr(varDatas(i)(10))  '调入门店
                    
                    '检查职员是否存在
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 3)
                    If lFItemID = 0 Then
                        sErr = "Employee Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillD01_Error
                    End If
                    .SetHeadValue "FEmpID", lFItemID
                    .SetHeadValue "FFManagerID", lFItemID
'                    .SetHeadValue "FManagerID", lFItemID
                    .SetHeadValue "FSManagerID", lFItemID
                                       

                End If
                

                '表体赋值

                .AddRow lRow
                
                
                 '检查物料是否存在
                 lFItemID = GetItemIDByNumber(CNN, CStr(varDatas(i)(5)), 4, lUnitID)
                 If lFItemID = 0 Then
                     sErr = "Material Master Data [" & CStr(varDatas(i)(5)) & "] is NOT created in ERP system."
                     GoTo SynBillD01_Error
                 End If
                .SetEntryValue "FItemID", lFItemID, lRow            '物料
                .SetEntryValue "FUnitID", lUnitID, lRow            '单位
                
                '数量必须大于0
                If Val(varDatas(i)(6)) <= 0 Then
                    sErr = "The quantity can NOT be zero."
                    GoTo SynBillD01_Error
                End If
                .SetEntryValue "Fauxqty", Val(varDatas(i)(6)), lRow            '数量
                .SetEntryValue "FAuxQtyMust", Val(varDatas(i)(6)), lRow          '应发数量
                .SetEntryValue "Fqty", Val(varDatas(i)(6)), lRow            '基本单位数量
                .SetEntryValue "FQtyMust", Val(varDatas(i)(6)), lRow          '基本单位应发数量
                

                '仓库仓位处理
                If Val(varDatas(i)(3)) = -1 Then
                    '当类型是调出时，调入仓库为T-WH，调出仓库为门店
                    .SetEntryValue "FDCStockID", lInTransitWH, lRow        '调入仓库
                            
                    '检查仓库是否存在
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 5)
                    If lFItemID = 0 Then
                         sErr = "Warehouse Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                         GoTo SynBillD01_Error
                     End If
                    .SetEntryValue "FSCStockID", lFItemID, lRow        '调出仓库
                    
'                         If Val(rsBillInfo.Fields("FIsStockMgr").Value) <> 0 And Val(rsBillInfo.Fields("FBinID").Value) = 0 Then
'                             sErr = "Bin Master Data [" & rsBillInfo.Fields("FBinCode").Value & "] is NOT created in ERP system or does not in the right bin group."
'                             GoTo SynBillD01_Error
'                         End If
'                        .SetEntryValue "FSCSPID", rsBillInfo.Fields("FBinID").Value, lRow        '调出仓位
                ElseIf Val(varDatas(i)(3)) = 1 Then
                    '当类型是调出时，调入仓库为门店，调出仓库为T-WH
                    .SetEntryValue "FSCStockID", lInTransitWH, lRow        '调出仓库
                            
                    '检查仓库是否存在
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 5)
                    If lFItemID = 0 Then
                         sErr = "Warehouse Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                         GoTo SynBillD01_Error
                     End If
                    .SetEntryValue "FDCStockID", lFItemID, lRow        '调入仓库
                    
'                         If Val(rsBillInfo.Fields("FIsStockMgr").Value) <> 0 And Val(rsBillInfo.Fields("FBinID").Value) = 0 Then
'                             sErr = "Bin Master Data [" & rsBillInfo.Fields("FBinCode").Value & "] is NOT created in ERP system or does not in the right bin group."
'                             GoTo SynBillD01_Error
'                         End If
'                        .SetEntryValue "FDCSPID", rsBillInfo.Fields("FBinID").Value, lRow        '调入仓位
                End If

                
                .SetEntryValue "FNote", CStr(varDatas(i)(8)), lRow          ''备注
                

                lRow = lRow + 1

            Next
            
'            CNN.Execute StrSqlTemp

            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
                checkBillData sDsn, lInterID, 41, 0, sErr, 16394 '审核单据
                
'                StrSqlTemp = "update T_t_Sales set FTransType=41,FInterID=" & lInterID & " where FUUID='" & strUUID & "'"
'                CNN.Execute StrSqlTemp
            End If
        End If
    End With

    Set clsBill = Nothing
    Set CNN = Nothing
    SynBillD01 = True
    GetObjectContext.SetComplete

   Exit Function

SynBillD01_Error:
    SynBillD01 = False
    GetObjectContext.SetAbort
    Set clsBill = Nothing
    Set CNN = Nothing
    sErr = "生成调拨单失败：" & sErr

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






