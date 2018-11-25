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


'/* ���۳��ⵥ Sales
Private Const CONST_B01_FBillNo As Long = 0 '���۵���

'Added by Nicky  - 20150504
'�����Ƿ����������ֶ�
Private Const CONST_B01_FType As Long = 1 '���۳��ⵥ����
'Added by Nicky  - 20150504

Private Const CONST_B01_FDATE As Long = 2 '���۳��ⵥ����
Private Const CONST_B01_FEntryID  As Long = 3 '����/�˻�
Private Const CONST_B01_FItemNumber  As Long = 4 '����
Private Const CONST_B01_FUOM As Long = 5 '���ϴ���
Private Const CONST_B01_FLotNumber As Long = 6 ' ����
Private Const CONST_B01_FProductionDate As Long = 7 'ԭ�ۣ���˰��
Private Const CONST_B01_FShelfLife As Long = 8 'ʵ�����۽���˰��
Private Const CONST_B01_FExpiryDate As Long = 9 'ԭ���۽���˰��=ԭ�ۣ���˰��*����
Private Const CONST_B01_FBoxBarcode As Long = 10 ' �ۿ۽�����˰��
Private Const CONST_B01_FQty As Long = 11 ' �ۿ۽���˰��
Private Const CONST_B01_FSPID As Long = 12 ' �ۿ���
Private Const CONST_B01_FReleaseDate As Long = 13 ' ����˰���۽��
Private Const CONST_B01_FCustomerCode As Long = 14 ' ˰��
Private Const CONST_B01_FCustomerName As Long = 15 ' ˰��
Private Const CONST_B01_FAddress As Long = 16 ' ����Ա
Private Const CONST_B01_FContact As Long = 17 ' ����Ա
Private Const CONST_B01_FPhone As Long = 18 ' ����Ա


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
' Purpose   : �������۳��ⵥ����� (POS TO ERP)
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
    Dim dblDiscountAmount As Double '�ۿ۶�
    Dim dblMarkDownAmt As Double ''�����ۿ۴���Ϊ PRWN �Ľ��
    Dim dblAmt As Double '���
    Dim dblDiscount As Double

    Dim lFItemID As Long
    Dim lUnitID As Long
    Dim dRate As Double

    Dim lSPGroupID As Long
    Dim blnIsSPGrop As Boolean
    Dim blnCheck As Boolean
    Dim sSourceBillNo As String
    Dim lROB As Long
    Dim blnEffective As Boolean ''�Ƿ������Ч����
    Dim isUpdateStore As Boolean

    Dim strMarkDown As String

    '���źͱ����ڹ���
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
    

    'ʹ��GUID��Ϊһ������ı�ʶ
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

                If i = 0 Then '��ͷ
                
                    '���POS Number�Ƿ���ڣ������������ظ�����
                    Set rs = CNN.Execute("select 1 from ICStockBill where FTranType=21 and isnull(FPOSNumber,'')='" & CStr(varDatas(i)(1)) & "' and FROB=" & CStr(varDatas(i)(2)))
                    If rs.RecordCount > 0 Then
                        sErr = "POS Number[" & CStr(varDatas(i)(1)) & "] has existed in ERP system."
                        GoTo SynBillB01_Error
                    End If
                                
                    .SetHeadValue "FDate", Format$(CDate(varDatas(i)(3)), "yyyy-mm-dd") '����
                    .SetHeadValue "FExplanation", Format(varDatas(i)(3), "HH:mm:ss") '����
                    .SetHeadValue "FPOSNumber", CStr(varDatas(i)(1))  'POS Number
                    
                    .SetHeadValue "FConsigneeAdd", CStr(varDatas(i)(15))  'Payment Method
                    
                    .SetHeadValue "FSaleStyle", 101 '���۷�ʽ:Ĭ������
                    .SetHeadValue "FMarketingStyle", 12530 '����ҵ������:Ĭ�����۳���
                    
                    .SetHeadValue "FROB", Val(varDatas(i)(2)) '�����ֱ�ʶ
                    lROB = Val(varDatas(i)(2))
                                    
                    '���ͻ��Ƿ����
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 1)
                    If lFItemID = 0 Then
                        sErr = "Customer Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillB01_Error
                    End If
                    .SetHeadValue "FSupplyID", lFItemID
                    '�жϿͻ��Ƿ������Զ����ɷ�Ʊ�Ŀͻ�
                    Set rs = CNN.Execute("select isnull(FAutoInv,0) FAutoInv from t_organization where FItemID=" & lFItemID)
                    bAutoInv = rs.Fields("FAutoInv").Value

                
                    '��鲿���Ƿ����
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 2)
                    If lFItemID = 0 Then
                        sErr = "Department Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillB01_Error
                    End If
                    .SetHeadValue "FDeptID", lFItemID
                    
                    '���ְԱ�Ƿ����
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 3)
                    If lFItemID = 0 Then
                        sErr = "Employee Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillB01_Error
                    End If
                    .SetHeadValue "FFManagerID", lFItemID
                    .SetHeadValue "FSManagerID", lFItemID
                
                End If
                
                
                '���帳ֵ

                .AddRow lRow
                
                ROW = lRow
                
                
                 '��������Ƿ����
                 lFItemID = GetItemIDByNumber(CNN, CStr(varDatas(i)(5)), 4, lUnitID)
                 If lFItemID = 0 Then
                     sErr = "Material Master Data [" & CStr(varDatas(i)(5)) & "] is NOT created in ERP system."
                     GoTo SynBillB01_Error
                 End If
                .SetEntryValue "FItemID", lFItemID, lRow            '����
                .SetEntryValue "FUnitID", lUnitID, lRow            '��λ
                
                '�����������0
                If Val(varDatas(i)(6)) <= 0 Then
                    sErr = "The quantity can NOT be zero."
                    GoTo SynBillB01_Error
                End If
                .SetEntryValue "Fauxqty", Val(varDatas(i)(6)) * lROB, lRow           '����
                .SetEntryValue "FAuxQtyMust", Val(varDatas(i)(6)) * lROB, lRow          'Ӧ������
                
                '���ֿ��Ƿ����
                lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 5)
                If lFItemID = 0 Then
                     sErr = "Warehouse Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                     GoTo SynBillB01_Error
                 End If
                .SetEntryValue "FDCStockID", lFItemID, lRow        '�����ֿ�
            
                .SetEntryValue "FConsignPrice", CNulls(Val(varDatas(i)(7)), 0), lRow           ''����
                .SetEntryValue "FDiscountAmount", CNulls(Val(varDatas(i)(12)), 0) * lROB, lRow          ''�ۿ۶�
                .SetEntryValue "FDiscountRate", (1 - CNulls(Val(varDatas(i)(11)), 0)) * 100, lRow       ''�ۿ���
                .SetEntryValue "FConsignAmount", CNulls(Val(varDatas(i)(13)), 0) * lROB, lRow          ''���۽��
                
                
                .SetEntryValue "FNote", CStr(varDatas(i)(14)), lRow          ''��ע
                
                lRow = lRow + 1
NextRow:
'                blnCheck = True ''�Ƿ������Ч����
            Next
            
            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
'                checkBillData sDsn, lInterID, 21, 0, sErr, 16394 '��˵���
                
                If checkBillData(sDsn, lInterID, 21, 0, sErr, 16394) = True Then  '��˵���
                    
                    '�Զ����ɹ�����Ʊ
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
    sErr = "Generate Sales Delivery Failed��" & sErr & Err.Description
    GetObjectContext.SetAbort
    Set clsBill = Nothing
    Set CNN = Nothing
    

End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateSI
' DateTime  : 2018-9-5
' Author    :
' Purpose   : ����ͻ��������Ϲ�ѡ���Զ����ɷ�Ʊ���������۷�Ʊ�����
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

            strSQL = "select FAccountID from t_Account where FNumber='1122.01'"
            Set rsTemp = ExecSQL(strSQL, sDsn)
            .SetHeadValue "FCussentAcctID", rsTemp.Fields("FAccountID").Value    '������Ŀ

        
            '���帳ֵ
            For j = 1 To rs.RecordCount
                
                            
                .AddRow lRow
                
                ROW = lRow

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


Private Function ExecSQL(ByVal ssql As String, ByVal dsn As String) As ADODB.Recordset
    Dim OBJ As Object
    Dim rs As ADODB.Recordset

    Set OBJ = CreateObject("BillDataAccess.GetData")
    Set rs = OBJ.ExecuteSQL(dsn, ssql)
    Set OBJ = Nothing
    Set ExecSQL = rs
    Set rs = Nothing
End Function





