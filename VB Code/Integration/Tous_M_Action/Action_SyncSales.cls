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
' Purpose   : �������۳��ⵥ�����
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
    
    '������
'    sDsn = "ConnectString={Persist Security Info=True;Provider=SQLOLEDB.1;User ID=sa;Password=sa;Data Source=.;Initial Catalog=Bluesea};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=BAD160C5-B4DB-4965-A8A7-0D373C885AE6"
'    strTablename = "t_POS_SAL"
    
    InitDataEnv sDsn
    Set CNN = datasource.Connection
    
    '��ȡ���ε����ļ��еĵ�����Ϣ�����Ϊ���򱨴�
    Set rsBill = ExecSQL("select distinct FPOSNumber,FSalesType from  " & strTablename & " order by FPOSNumber", sDsn)
    If rsBill.RecordCount <= 0 Then
        sErr = "There is no data in the EDI file, please contact the administrator for double checking."
        GoTo SynBillB01_Error
    End If
    
    '��ȡ������ϸ��׼�����е���
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
            
                '���POS Number�Ƿ���ڣ������������ظ�����
                Set rs = ExecSQL("select 1 from ICStockBill where FTranType=21 and isnull(FPOSNumber,'')='" & rsBill.Fields("FPOSNumber") & "' and FROB=" & rsBill.Fields("FSalesType"), sDsn)
                If rs.RecordCount > 0 Then
                    sErr = "POS Number[" & rsBill.Fields("FPOSNumber").Value & "] has existed in ERP system."
                    GoTo SynBillB01_Error
                End If
                            
                .SetHeadValue "FDate", CDate(rsBillInfo.Fields("FDate").Value)  '����
                .SetHeadValue "FPOSNumber", CStr(rsBillInfo.Fields("FPOSNumber").Value)  'POS Number
                
                .SetHeadValue "FSaleStyle", 101 '���۷�ʽ:Ĭ������
                .SetHeadValue "FMarketingStyle", 12530 '����ҵ������:Ĭ�����۳���
                
                .SetHeadValue "FROB", Val(rsBill.Fields("FSalesType")) '�����ֱ�ʶ
                lROB = Val(rsBill.Fields("FSalesType"))
                
                '���ͻ��Ƿ����
                If Val(rsBillInfo.Fields("FCustID").Value) = 0 Then
                    sErr = "Customer Master Data [" & rsBillInfo.Fields("FStoreCode").Value & "] is NOT created in ERP system."
                    GoTo SynBillB01_Error
                End If
                .SetHeadValue "FSupplyID", rsBillInfo.Fields("FCustID").Value
            
                '��鲿���Ƿ����
                If Val(rsBillInfo.Fields("FDeptID").Value) = 0 Then
                    sErr = "Department Master Data [" & rsBillInfo.Fields("FStoreCode").Value & "] is NOT created in ERP system."
                    GoTo SynBillB01_Error
                End If
                .SetHeadValue "FDeptID", rsBillInfo.Fields("FDeptID")
                
                '���ְԱ�Ƿ����
                If Val(rsBillInfo.Fields("FEmpID").Value) = 0 Then
                    sErr = "Employee Master Data [" & rsBillInfo.Fields("FStoreCode").Value & "] is NOT created in ERP system."
                    GoTo SynBillB01_Error
                End If
                .SetHeadValue "FFManagerID", rsBillInfo.Fields("FEmpID")
                .SetHeadValue "FSManagerID", rsBillInfo.Fields("FEmpID")

            
                '���帳ֵ
                For j = 1 To rsBillInfo.RecordCount
                                
                    .AddRow lRow
                    
                    
                    Row = lRow
    
                     '��������Ƿ����
                     If Val(rsBillInfo.Fields("FItemID").Value) = 0 Then
                         sErr = "Material Master Data [" & rsBillInfo.Fields("FSKU").Value & "] is NOT created in ERP system."
                         GoTo SynBillB01_Error
                     End If
                    .SetEntryValue "FItemID", rsBillInfo.Fields("FItemID").Value, lRow            '����
                    .SetEntryValue "FUnitID", rsBillInfo.Fields("FUnitID").Value, lRow            '��λ
                    
                    '�����������0
                    If Val(rsBillInfo.Fields("FQty").Value) <= 0 Then
                        sErr = "The quantity can NOT be zero."
                        GoTo SynBillB01_Error
                    End If
                    .SetEntryValue "Fauxqty", Val(rsBillInfo.Fields("FQty").Value) * lROB, lRow           '����
                    .SetEntryValue "FAuxQtyMust", Val(rsBillInfo.Fields("FQty").Value) * lROB, lRow          'Ӧ������
                    
                     If Val(rsBillInfo.Fields("FStockID").Value) = 0 Then
                         sErr = "Warehouse Master Data [" & rsBillInfo.Fields("FStoreCode").Value & "] is NOT created in ERP system."
                         GoTo SynBillB01_Error
                     End If
                    .SetEntryValue "FDCStockID", rsBillInfo.Fields("FStockID").Value, lRow        '�����ֿ�
                
                    .SetEntryValue "FConsignPrice", CNulls(rsBillInfo.Fields("FPrice").Value, 0), lRow           ''����
                    .SetEntryValue "FDiscountAmount", CNulls(rsBillInfo.Fields("FDiscountAmount").Value, 0) * lROB, lRow          ''�ۿ۶�
                    .SetEntryValue "FDiscountRate", CNulls(rsBillInfo.Fields("FDiscountRate").Value, 0), lRow           ''�ۿ���
                    .SetEntryValue "FConsignAmount", CNulls(rsBillInfo.Fields("FTotalAmount").Value, 0) * lROB, lRow          ''���۽��
                    
                    
                    .SetEntryValue "FNote", rsBillInfo.Fields("FNote").Value, lRow          ''��ע
                

                '=========Դ����Ϣ===========
'                .SetEntryValue "FSourceTranType", "81", lRow          ''Դ������
'                .SetEntryValue "FSourceBillNo", rsBillInfo.Fields("FBillNo").Value, lRow          ''Դ�����
'                .SetEntryValue "FSourceInterId", rsBillInfo.Fields("FInterID").Value, lRow          ''Դ������
'                .SetEntryValue "FSourceEntryID", rsBillInfo.Fields("FEntryID").Value, lRow          ''Դ���к�
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
                    If checkBillData(sDsn, lInterID, 21, 0, sErr, 16394) = True And bAutoInv = True Then '��˵���
                        
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
    sErr = "Generate Sales Delivery failed��" & sErr & Err.Description
    
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
' Purpose   : ����ͻ��������Ϲ�ѡ���Զ����ɷ�Ʊ���������۷�Ʊ�����
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
    
    '��ȡ���ⵥ��ϸ��׼�����ɷ�Ʊ
    strSQL = "select t1.FInterID,t1.FROB,t1.FBillNo,t1.FPOSNumber,t1.FSupplyID FCustID,t1.FDate,t1.FSaleStyle,t1.FDeptID,t1.FFManagerID FEmpID,"
    strSQL = strSQL & vbCrLf & "t2.FEntryID, t2.fItemID , t2.FUnitID, t2.FConsignPrice, t2.FQty, t2.FDiscountRate, t2.FDiscountAmount, t2.FConsignAmount,"
    strSQL = strSQL & vbCrLf & "t2.FNote, t3.FTaxRate, t3.FPriceDecimal from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
    strSQL = strSQL & vbCrLf & "inner join t_ICItem t3 on t2.FItemID=t3.FItemID where t1.FInterID=" & CStr(lInterID)
    '�޵�����Ʒ��������Ʒ������Ʊ
    strSQL = strSQL & vbCrLf & "and t2.FConsignAmount<>0"
    strSQL = strSQL & vbCrLf & "order by t1.FInterID,t2.FEntryID"
    
    Set rs = ExecSQL(strSQL, sDsn)
            
    Set clsBill = New AutoOldBill
    
    lRow = 1
    
    With clsBill
        .sDsn = sDsn
        
        bRet = .NewBill(80)

        If bRet = True Then
                        
            .SetHeadValue "FCustID", rs.Fields("FCustID").Value '�ͻ�
            .SetHeadValue "FDeptID", rs.Fields("FDeptID") '����
            .SetHeadValue "FEmpID", rs.Fields("FEmpID") 'ҵ��Ա
            .SetHeadValue "FDate", CDate(rs.Fields("FDate").Value)  '����
            .SetHeadValue "FPOSNumber", CStr(rs.Fields("FPOSNumber").Value)  'POS Number
            .SetHeadValue "FSaleStyle", rs.Fields("FSaleStyle").Value '���۷�ʽ
'            .SetHeadValue "FMarketingStyle", rs.Fields("FSaleStyle").Value '����ҵ������:Ĭ�����۳���
            
            .SetHeadValue "FROB", Val(rs.Fields("FROB")) '�����ֱ�ʶ
            lROB = Val(rs.Fields("FROB"))
            
            .SetHeadValue "FCurrencyID", 1   '�ұ�Ĭ�ϱ�λ��
            .SetHeadValue "FExchangeRateType", 1    '�������ͣ�Ĭ�Ϲ�˾����
            .SetHeadValue "FExchangeRate", 1    '���ʣ�Ĭ��1

        
            '���帳ֵ
            For j = 1 To rs.RecordCount
                
                            
                .AddRow lRow
                
                Row = lRow

                .SetEntryValue "FItemID", rs.Fields("FItemID").Value, lRow            '����
                .SetEntryValue "FUnitID", rs.Fields("FUnitID").Value, lRow            '��λ
                
                .SetEntryValue "Fauxqty", Val(rs.Fields("FQty").Value) * lROB, lRow           '����
                .SetEntryValue "Fqty", Val(rs.Fields("FQty").Value) * lROB, lRow          '������λ����
            
                .SetEntryValue "FAuxTaxPrice", Val(rs.Fields("FConsignPrice").Value), lRow           '��˰����
                .SetEntryValue "FAuxPriceDiscount", Val(rs.Fields("FConsignAmount").Value) / Val(rs.Fields("FQty").Value), lRow         'ʵ�ʺ�˰����
                .SetEntryValue "Fauxprice", KDRound((Val(rs.Fields("FConsignPrice").Value) / (1 + Val(rs.Fields("FTaxRate").Value) / 100)), Val(rs.Fields("FPriceDecimal").Value)), lRow     '��˰����
                
                .SetEntryValue "FAmtDiscount", Val(rs.Fields("FDiscountAmount").Value) * lROB, lRow          ''�ۿ۶�
                .SetEntryValue "FStdAmtDiscount", Val(rs.Fields("FDiscountAmount").Value) * lROB, lRow          ''�ۿ۶��λ�ң�
                .SetEntryValue "FDiscountRate", Val(rs.Fields("FDiscountRate").Value), lRow           ''�ۿ���
                
                .SetEntryValue "FTaxRate", Val(rs.Fields("FTaxRate").Value), lRow         '˰��
                .SetEntryValue "FTaxAmount", KDRound(Val(rs.Fields("FConsignAmount").Value) - (Val(rs.Fields("FConsignAmount").Value) / (1 + Val(rs.Fields("FTaxRate").Value) / 100)), 2) * lROB, lRow        '˰��
                .SetEntryValue "FStdTaxAmount", KDRound(Val(rs.Fields("FConsignAmount").Value) - (Val(rs.Fields("FConsignAmount").Value) / (1 + Val(rs.Fields("FTaxRate").Value) / 100)), 2) * lROB, lRow       '˰���λ�ң�
                
                .SetEntryValue "FAllAmount", Val(rs.Fields("FConsignAmount").Value) * lROB, lRow          '��˰�ϼ�
                .SetEntryValue "FStdAllAmount", Val(rs.Fields("FConsignAmount").Value) * lROB, lRow       '��˰�ϼƣ���λ�ң�
                .SetEntryValue "FAmount", KDRound((Val(rs.Fields("FConsignAmount").Value) / (1 + Val(rs.Fields("FTaxRate").Value) / 100)), 2) * lROB, lRow         '���
                .SetEntryValue "FStdAmount", KDRound((Val(rs.Fields("FConsignAmount").Value) / (1 + Val(rs.Fields("FTaxRate").Value) / 100)), 2) * lROB, lRow        '����λ�ң�
                
                .SetEntryValue "FNote", rs.Fields("FNote").Value, lRow          ''��ע
            

            '=========Դ����Ϣ===========
                .SetEntryValue "FSourceTranType", "21", lRow          ''Դ������
                .SetEntryValue "FSourceBillNo", rs.Fields("FBillNo").Value, lRow          ''Դ�����
                .SetEntryValue "FSourceInterId", rs.Fields("FInterID").Value, lRow          ''Դ������
                .SetEntryValue "FSourceEntryID", rs.Fields("FEntryID").Value, lRow          ''Դ���к�
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
                checkBillData sDsn, lInterID, 80, 0, sErr, 16394 '��˵���
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
    sErr = "Generate Sales Invoice error��" & sErr & Err.Description
    
End Function

