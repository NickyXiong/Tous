VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "Action_SyncStockAdjustment"
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
' Procedure : SyncPurchaseReceipt
' DateTime  : 2018-9-5
' Author    :
' Purpose   : 生成外购入库单并审核
'---------------------------------------------------------------------------------------
'
Public Function SyncStockAdjustment(ByVal sDsn As String, strTablename As String, ByRef strInterIDs As String, ByRef sBillNOs As String, ByRef sErr As String) As Boolean
 
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
    
On Error GoTo SynBillA01_Error
    
    '测试用
    sDsn = "ConnectString={Persist Security Info=True;Provider=SQLOLEDB.1;User ID=sa;Password=sa;Data Source=.;Initial Catalog=Bluesea};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=BAD160C5-B4DB-4965-A8A7-0D373C885AE6"
    strTablename = "t_POS_pur"
    
    InitDataEnv sDsn
    Set CNN = datasource.Connection
    
    '获取本次导入文件中的单据信息，如果为空则报错
    Set rsBill = ExecSQL("select distinct FPOSNumber from  " & strTablename & " order by FPOSNumber", sDsn)
    If rsBill.RecordCount <= 0 Then
        sErr = "There is no data in the EDI file, please contact the administrator for double checking."
        GoTo SynBillA01_Error
    End If
    
    '获取导入明细，准备进行导入
    strSQL = "select t1.*,t2.FSupplyID,isnull(t3.FItemID,0) FItemID,isnull(t3.FUnitID,0) FUnitID,isnull(t5.FItemID,0) FEmpID,isnull(t6.FItemID,0) FStockID,"
    strSQL = strSQL & vbCrLf & "t6.FIsStockMgr,isnull(t7.FSPID,0) FBinID,t2.FBillNo,t2.FInterID,t3.FEntryID "
    strSQL = strSQL & vbCrLf & "from t_POS_PUR t1 inner join POOrder t2 on t1.fponumber=t2.FBillNo"
    strSQL = strSQL & vbCrLf & "left join (select k2.FNumber,k1.* from POOrderEntry k1 inner join t_ICItem k2 on k1.FItemID=k2.FItemID)t3"
    strSQL = strSQL & vbCrLf & "on t2.FInterID=t3.FInterID and t3.FNumber=t1.fSKU"
    strSQL = strSQL & vbCrLf & "left join t_Emp t5 on t1.FStoreCode=t5.FNumber"
    strSQL = strSQL & vbCrLf & "left join t_Stock t6 on t1.FStoreCode=t6.FNumber"
    strSQL = strSQL & vbCrLf & "left join t_StockPlace t7 on t1.FBinCode=t7.FNumber and t6.FSPGroupID=t7.FSPGroupID"
    strSQL = strSQL & vbCrLf & "order by t1.FPOSNumber,t1.FStoreCode"
    
    Set rsBillInfo = ExecSQL(strSQL, sDsn)
            
    Set clsBill = New AutoOldBill
    
    lRow = 1
    
    With clsBill
        .sDsn = sDsn
                
        For i = 1 To rsBill.RecordCount
        
            bRet = .NewBill(1)
    
            If bRet = True Then
            
                rsBillInfo.Filter = "FPOSNumber='" & rsBill.Fields("FPOSNumber") & "'"
            
                '检查POS Number是否存在，存在则不允许重复导入
                Set rs = ExecSQL("select 1 from ICStockBill where FTranType=1 and isnull(FPOSNumber,'')='" & rsBill.Fields("FPOSNumber") & "'", sDsn)
                If rs.RecordCount > 0 Then
                    sErr = "POS Number[" & rsBill.Fields("FPOSNumber").Value & "] has existed in ERP system."
                    GoTo SynBillA01_Error
                End If
                            
                .SetHeadValue "FDate", Format(Now, "yyyy-mm-dd") '日期
                .SetHeadValue "FSupplyID", rsBillInfo.Fields("FSupplyID").Value '供应商默认取采购订单供应商
                .SetHeadValue "FPOMode", 36680 '采购模式:  默认为普通采购
                .SetHeadValue "FPOStyle", 252  '采购方式: 默认为赊购
                
                .SetHeadValue "FEmpID", rsBillInfo.Fields("FEmpID")
                .SetHeadValue "FFManagerID", rsBillInfo.Fields("FEmpID")
                .SetHeadValue "FManagerID", rsBillInfo.Fields("FEmpID")
                .SetHeadValue "FSManagerID", rsBillInfo.Fields("FEmpID")
                
            
                '表体赋值
                For j = 1 To rsBillInfo.RecordCount
                                
                    .AddRow lRow
                    
                    Row = lRow
    
                     '检查物料是否存在
                     If Val(rsBillInfo.Fields("FItemID").Value) = 0 Then
                         sErr = "Material Master Data [" & rsBillInfo.Fields("FSKU").Value & "] is NOT created in ERP system."
                         GoTo SynBillA01_Error
                     End If
                    .SetEntryValue "FItemID", rsBillInfo.Fields("FItemID").Value, lRow            '物料
                    .SetEntryValue "FUnitID", rsBillInfo.Fields("FUnitID").Value, lRow            '单位
                    
                    '数量必须大于0
                    If Val(rsBillInfo.Fields("FQty").Value) <= 0 Then
                        sErr = "The quantity can NOT be zero."
                        GoTo SynBillA01_Error
                    End If
                    .SetEntryValue "Fauxqty", Val(rsBillInfo.Fields("FQty").Value) * lROB, lRow           '数量
                    .SetEntryValue "FAuxQtyMust", Val(rsBillInfo.Fields("FQty").Value) * lROB, lRow          '应收数量
                    
                     If Val(rsBillInfo.Fields("FStockID").Value) = 0 Then
                         sErr = "Warehouse Master Data [" & rsBillInfo.Fields("FStoreCode").Value & "] is NOT created in ERP system."
                         GoTo SynBillA01_Error
                     End If
                    .SetEntryValue "FDCStockID", rsBillInfo.Fields("FStockID").Value, lRow        '收货仓库
                    
                     If Val(rsBillInfo.Fields("FIsStockMgr").Value) <> 0 And Val(rsBillInfo.Fields("FBinID").Value) = 0 Then
                         sErr = "Bin Master Data [" & rsBillInfo.Fields("FBinCode").Value & "] is NOT created in ERP system or does not in the right bin group."
                         GoTo SynBillA01_Error
                     End If
                    .SetEntryValue "FDCSPID", rsBillInfo.Fields("FBinID").Value, lRow        '收货仓位

                     '=========源单信息===========
                     .SetEntryValue "FSourceTranType", 71, lRow          ''源单类型
                     .SetEntryValue "FSourceBillNo", rsBillInfo.Fields("FBillNo").Value, lRow          ''源单编号
                     .SetEntryValue "FSourceInterId", rsBillInfo.Fields("FInterID").Value, lRow          ''源单内码
                     .SetEntryValue "FSourceEntryID", rsBillInfo.Fields("FEntryID").Value, lRow          ''源单行号
                     .SetEntryValue "FOrderInterID", rsBillInfo.Fields("FInterID").Value, lRow            '
                     .SetEntryValue "FOrderEntryID", rsBillInfo.Fields("FEntryID").Value, lRow           '
                     .SetEntryValue "FOrderBillNo", rsBillInfo.Fields("FBillNo").Value, lRow
                    '=============================
                    
                    .SetEntryValue "FNote", rsBillInfo.Fields("FNote").Value, lRow          ''备注
               
                    lRow = lRow + 1
NextRow:
                    rsBillInfo.MoveNext
                Next j
NewRow:
                lInterID = 0
                strBillNo = ""
                bRet = .SaveBill(lInterID, strBillNo, sErr)
                If bRet = True Then
                    checkBillData sDsn, lInterID, 1, 0, sErr, 16394 '审核单据
      
                    strInterIDs = strInterIDs & CStr(lInterID) & ","
                    sBillNOs = sBillNOs & strBillNo & ","

                End If
                
                rsBill.MoveNext
                rsBillInfo.Filter = ""
            
            End If
        Next i
            
            
    End With
    
    Set clsBill = Nothing
    Set CNN = Nothing
    SyncStockAdjustment = True
    GetObjectContext.SetComplete

    Exit Function

SynBillA01_Error:
    SyncStockAdjustment = False
'    GetObjectContext.SetAbort
    Set clsBill = Nothing
    Set CNN = Nothing
    sErr = "Generate Purchase Receipt failed：" & sErr & Err.Description
    
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



