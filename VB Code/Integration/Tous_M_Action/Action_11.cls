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
Public Function SynBillB01(ByVal sDsn As String, varDatas As Variant, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String, ByRef strUUID As String) As Boolean
    Dim cnn As ADODB.Connection
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

    Dim strBarCode As String
    Dim strMarkDown As String

    '���źͱ����ڹ���
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
            
                If UBound(varDatas(i)) <> 18 Then
                    sErr = "�����ļ���ʽ��������ϵϵͳ����Ա"
                    Exit Function
                End If

                If i = 0 Then '��ͷ

                    sSourceBillNo = Trim(CStr(varDatas(i)(CONST_B01_FBillNo)))

                    If IsExitBill(cnn, sSourceBillNo, "SEOrder", 81, lSOInterID, "FBillNo") = False Then
                        sErr = "���۶��� [" & sSourceBillNo & "] ������"
                        Exit Function
                    Else
                        StrSql = "select t3.FBarcode,t1.FCustID,t1.FDeptID,t1.FEmpID,t1.FMangerID,t2.FEntryID,t2.FItemID,t2.FAuxTaxPrice,FTaxAmount FDiscountAmt,t2.FTaxRate FDiscountRate,FAuxStockQty,(case when day(t1.FHeadSelfS0154)>25 then 102 else 101 end) FSaleStyle,t1.FHeadSelfS0156 SalesMode"
                        StrSql = StrSql & vbCrLf & "from SEOrder t1 inner join SEOrderEntry t2 on t1.FInterID=t2.FInterID"
                        StrSql = StrSql & vbCrLf & "inner join t_icitem t3 on t2.FItemID=t3.FItemID"
                        StrSql = StrSql & vbCrLf & "where t1.FInterID =" & lSOInterID
                        Set rsTemp = cnn.Execute(StrSql)
                    End If
                
                    .SetHeadValue "FMarketingStyle", 12530 '����ҵ������ ��Ĭ�����۳�������
                    
                    If Len(CStr(varDatas(i)(CONST_B01_FDATE))) <= 0 Then
                        sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:���ڲ���Ϊ��"
                        Exit Function
                    End If
                    .SetHeadValue "FHeadSelfB0157", rsTemp.Fields("SalesMode").Value   '��������
                    .SetHeadValue "FDate", Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd")  ' ����
                    .SetHeadValue "FSettleDate", Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd")  ' �տ�����


                    .SetHeadValue "FDeptID", rsTemp.Fields("FDeptID").Value   '����
                    .SetHeadValue "FSupplyID", rsTemp.Fields("FCustID").Value  '������λ

                    '���۷�ʽ�� ����
                    .SetHeadValue "FSaleStyle", 101


                    'ְԱ�ֶ�ȫ��Ĭ��Ϊ 001
                    lFItemID = GetItemID(cnn, "001", 3)
                    If lFItemID = 0 Then
                        sErr = "Can NOT find Employee [001]"
                        Exit Function
                    End If
                    .SetHeadValue "FFManagerID", lFItemID '������¼
                    .SetHeadValue "FEmpID", rsTemp.Fields("FEmpID").Value   'ҵ��Ա"
                    .SetHeadValue "FSManagerID", lFItemID   '����
                    .SetHeadValue "FManagerID", lFItemID  '����


                    .SetHeadValue "FRob", lROB '������
                End If
                
                
                rsTemp.Filter = "FEntryID=" & Trim(CStr(varDatas(i)(CONST_B01_FEntryID)))
                If Not rsTemp Is Nothing Then
                    If rsTemp.RecordCount > 0 Then
                    Else
'                        sErr = "��Ŀ�� [" & sSourceBillNo & "] ������"
                        sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:��Ŀ�Ų�����"
                        Exit Function
                    End If
                End If

                '���帳ֵ

                .AddRow lRow
                
                '��ͬ�������ŵ����յ������ϲ� modify 20160316
                If i <> 0 And strProduct = Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) And strBatch = CStr(Trim(varDatas(i)(CONST_B01_FLotNumber))) _
                And strPeriodDate = Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate))) And EntryID = Trim(CStr(varDatas(i)(CONST_B01_FEntryID))) Then
                    QTY = QTY + Val(varDatas(i)(CONST_B01_FQty)) * lROB
                    .SetEntryValue "Fauxqty", QTY, Row         '����
                    .SetEntryValue "FAuxQtyMust", (QTY - Val(rsTemp.Fields("FAuxStockQty").Value)) * lROB, Row
                    
                    StrSqlTemp = StrSqlTemp & vbCrLf & "insert T_t_Sales  values('���۶���','" & sSourceBillNo & _
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
'                    sErr = "�ֿ�[C004]������"
                    sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:�ֿ�[C014]������"
                    Exit Function
                End If
                .SetEntryValue "FDCStockID", lFItemID, lRow            '�����ֿ�
                lStockID = lFItemID
                
                EntryID = Trim(CStr(varDatas(i)(CONST_B01_FEntryID)))
                
                blnIsSPGrop = GetIsDCSP(cnn, lStockID, lSPGroupID)
                If blnIsSPGrop = True Then
                    '��λ����ݲֿ����Բ�λ����Ҳ�λ��Ϣ
                    lFItemID = GetDCSPID(cnn, Trim(CStr(varDatas(i)(CONST_B01_FSPID))), lSPGroupID)
                    If lFItemID = 0 Then
'                        sErr = "��λ[" & Trim(CStr(varDatas(i)(CONST_B01_FSPID))) & "] ������"
                        sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:��λ[" & Trim(CStr(varDatas(i)(CONST_B01_FSPID))) & "] ������"
                        Exit Function
                    End If
                    lSPID = lFItemID
                    .SetEntryValue "FDCSPID", lFItemID, lRow              '��λ
                End If

                lFItemID = GetItemID(cnn, Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))), 4, lUnitID, dRate)
                If lFItemID = 0 Then
'                    sErr = "������Ʒ[" & Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) & "]������"
                    sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:������Ʒ[" & Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) & "]������"
                    Exit Function
                End If
                
                If Val(rsTemp.Fields("FItemID").Value) <> lFItemID Then
'                    sErr = "������Ʒ[" & Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) & "]�����۶�������Ʒ��һ�£��������ݻ���Ŀ���Ƿ���ȷ"
                    sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:������Ʒ[" & Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) & "]�����۶�������Ʒ��һ�£��������ݻ���Ŀ���Ƿ���ȷ"
                    Exit Function
                End If
                
                .SetEntryValue "FItemID", lFItemID, lRow     '����
                .SetEntryValue "FUnitID", lUnitID, lRow      '��λ
                
                strProduct = Trim(CStr(varDatas(i)(CONST_B01_FItemNumber)))
                strBarCode = Trim(CStr(varDatas(i)(CONST_B01_FBoxBarcode)))

                '���źͱ�����
                If Len(CStr(varDatas(i)(CONST_B01_FLotNumber))) <= 0 Then
'                    sErr = "���۶��� [" & sSourceBillNo & "] :���Ų���Ϊ��"
                    sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:���Ų���Ϊ��"
                    Exit Function
                End If
                .SetEntryValue "FBatchNo", CStr(Trim(varDatas(i)(CONST_B01_FLotNumber))), lRow              ''�ֿ�
                strBatch = CStr(Trim(varDatas(i)(CONST_B01_FLotNumber)))
                
                If Len(CStr(varDatas(i)(CONST_B01_FProductionDate))) <= 0 Then
'                    sErr = "���۶��� [" & sSourceBillNo & "] :�������ڲ���Ϊ��"
                    sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:�������ڲ���Ϊ��"
                    Exit Function
                End If
'                .SetEntryValue "FKFDate", CStr(Trim(varDatas(i)(CONST_B01_FProductionDate))), lRow              ''�ֿ�
                .SetEntryValue "FKFDate", DateAdd("d", -Val(varDatas(i)(CONST_B01_FShelfLife)), Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate)))), lRow            ''��������
                
                If Len(CStr(varDatas(i)(CONST_B01_FShelfLife))) <= 0 Then
'                    sErr = "���۶��� [" & sSourceBillNo & "] :�����ڲ���Ϊ��"
                    sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:�����ڲ���Ϊ��"
                    Exit Function
                End If
                .SetEntryValue "FKFPeriod", Val(varDatas(i)(CONST_B01_FShelfLife)), lRow              ''�ֿ�
                
                If Len(CStr(varDatas(i)(CONST_B01_FExpiryDate))) <= 0 Then
'                    sErr = "���۶��� [" & sSourceBillNo & "] :�������ڲ���Ϊ��"
                    sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:�������ڲ���Ϊ��"
                    Exit Function
                End If
                .SetEntryValue "FPeriodDate", Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate))), lRow '��Ч����
                strPeriodDate = Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate)))

                If Val(CStr(varDatas(i)(CONST_B01_FQty))) <= 0 Then
'                    sErr = "���۶��� [" & sSourceBillNo & "] :������������Ϊ����"
                    sErr = "���۶���[" & sSourceBillNo & "]������Ϣ:������������Ϊ����"
                    Exit Function
                End If
                .SetEntryValue "Fauxqty", Val(varDatas(i)(CONST_B01_FQty)) * lROB, lRow           '����
                .SetEntryValue "FAuxQtyMust", (Val(varDatas(i)(CONST_B01_FQty)) - Val(rsTemp.Fields("FAuxStockQty").Value)) * lROB, lRow       'Ӧ������
                QTY = Val(varDatas(i)(CONST_B01_FQty)) * lROB
                
                '����
                .SetEntryValue "FConsignPrice", Val(rsTemp.Fields("FAuxTaxPrice").Value), lRow
                '���۽��
                .SetEntryValue "FConsignAmount", Val(varDatas(i)(CONST_B01_FQty)) * _
                                Val(rsTemp.Fields("FAuxTaxPrice").Value) * lROB * _
                                (1 - Val(rsTemp.Fields("FDiscountRate").Value) / 100), lRow
                
                '�ۿ���
                .SetEntryValue "FDiscountRate", Val(rsTemp.Fields("FDiscountRate").Value), lRow
                '��˰�ۿ۶�
'                .SetEntryValue "FDiscountAmount", Val(varDatas(i)(CONST_B01_FQty)) * _
'                                Val(rsTemp.Fields("FAuxTaxPrice").Value) * lROB * _
'                                (Val(rsTemp.Fields("FDiscountRate").Value) / 100), lRow
                .SetEntryValue "FDiscountAmount", Val(rsTemp.Fields("FDiscountAmt").Value), lRow
                
                ''������ϵ�ֶ�
                .SetEntryValue "FSourceTranType", 81, lRow           '
                .SetEntryValue "FSourceInterID", lSOInterID, lRow            '
                .SetEntryValue "FSourceBillNo", sSourceBillNo, lRow
                .SetEntryValue "FSourceEntryID", CNulls(rsTemp.Fields("FEntryID"), 0), lRow '
                .SetEntryValue "FOrderInterID", lSOInterID, lRow           '
                .SetEntryValue "FOrderEntryID", CNulls(rsTemp.Fields("FEntryID"), 0), lRow           '
                .SetEntryValue "FOrderBillNo", sSourceBillNo, lRow
                
                StrSqlTemp = StrSqlTemp & vbCrLf & "insert T_t_Sales  values('���۶���','" & sSourceBillNo & _
                                    "'," & lStockID & "," & lSPID & ",'" & Trim(CStr(varDatas(i)(CONST_B01_FBoxBarcode))) & _
                                    "',1,'" & Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd") & _
                                    "','" & Format(varDatas(i)(CONST_B01_FDATE), "yyyymmddhhmmss") & _
                                    "',0,0,'" & strUUID & "',1," & lStockID & "," & lSPID & ",'" & Format(Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate))), "yyyy-mm") & "')"
                
                lRow = lRow + 1
NextRow:
'                blnCheck = True ''�Ƿ������Ч����
            Next
            
            cnn.Execute StrSqlTemp

            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
'                checkBillData sDsn, lInterID, 21, 0, sErr, 16394 '��˵���
                
                StrSqlTemp = "update T_t_Sales set FTransType=21,FInterID=" & lInterID & " where FUUID='" & strUUID & "'"
                cnn.Execute StrSqlTemp
            End If
        End If
    End With
    
'    '=====���������ݿ�======
'
'     '��ȡ�����ݿ������ַ���
'    strDSN2 = TransfersDsn("st20160407", sDsn)
'
'    '��ѯ���ε�������ݰ����������·��鵼��
'    ssql = "select isnull(FYearMonth,'') FYearMonth from T_t_sales where FUUID='" & strUUID & "' and isnull(FYearMonth,'') <>'' group by FYearMonth"
'    Set rsYearMonth = ExecSQL(ssql, sDsn)
'
'
'    '���ݵ����ռ���Ӧ���������ݿ����Ƿ����,���������½�
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
'        '�����ε������ݰ��������·��鵼���Ӧ�����ݿ����
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
    sErr = "�������۳��ⵥʧ�ܣ�" & sErr & Err.Description
    

End Function


'---------------------------------------------------------------------------------------
' Procedure : SynBillD01
' DateTime  : 2013-1-25 00:05
' Author    :
' Purpose   : ���ɵ����������
'---------------------------------------------------------------------------------------
'
Public Function SynBillD01(ByVal sDsn As String, varDatas As Variant, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String) As Boolean
    Dim cnn As ADODB.Connection
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
    

    'ʹ��GUID��Ϊһ������ı�ʶ
    Set objTypeLib = CreateObject("Scriptlet.TypeLib")
    strUUID = CStr(objTypeLib.Guid)
    strUUID = Mid(strUUID, 1, InStr(1, strUUID, "}"))
    

    StrSqlTemp = ""
    
    With clsBill
        .sDsn = sDsn

        bRet = .NewBill(41)

        If bRet = True Then
            For i = LBound(varDatas) To UBound(varDatas)
            


                If i = 0 Then '��ͷ

                    sSourceBillNo = Trim(CStr(varDatas(i)(CONST_B01_FBillNo)))

                    If IsExitBill(cnn, sSourceBillNo, "ICSTOCKBILL", 41, , "FBillNo") = True Then
                        sErr = "������[" & sSourceBillNo & "]�����ɣ��������ظ�����"
                        Exit Function
                    End If
                    .SetHeadValue "FBillNo", sSourceBillNo  ' ����

                    If Len(CStr(varDatas(i)(CONST_B01_FDATE))) <= 0 Then
                        sErr = "������[" & sSourceBillNo & "] :���ڲ���Ϊ��"
                        Exit Function
                    End If
                    .SetHeadValue "FDate", Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd")  ' ����
'                    .SetHeadValue "FSettleDate", Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd")  ' �տ�����


                    'ְԱ�ֶ�ȫ��Ĭ��Ϊ 001
                    lFItemID = GetItemID(cnn, "001", 3)
                    If lFItemID = 0 Then
                        sErr = "ְԱ[001]������"
                        Exit Function
                    End If
                    .SetHeadValue "FFManagerID", lFItemID '������¼
                    .SetHeadValue "FEmpID", lFItemID   'ҵ��Ա"
                    .SetHeadValue "FSManagerID", lFItemID   '����
                    .SetHeadValue "FManagerID", lFItemID  '����

                    .SetHeadValue "FRob", lROB '������

                End If
                

                '���帳ֵ

                .AddRow lRow
                
                
                
                lFItemID = GetItemID(cnn, "C014", 5)
                If lFItemID = 0 Then
                    sErr = "������[" & sSourceBillNo & "] :�����ֿ�[C014] ������"
                    Exit Function
                End If
                .SetEntryValue "FSCStockID", lFItemID, lRow            '�����ֿ�
                lStockID = lFItemID
                
                
                blnIsSPGrop = GetIsDCSP(cnn, lStockID, lSPGroupID)
                If blnIsSPGrop = True Then
                    '��λ����ݲֿ����Բ�λ����Ҳ�λ��Ϣ
                    lFItemID = GetDCSPID(cnn, Trim(CStr(varDatas(i)(CONST_B01_FSPID))), lSPGroupID)
                    If lFItemID = 0 Then
                        sErr = "������[" & sSourceBillNo & "] :������λ[" & Trim(CStr(varDatas(i)(CONST_B01_FSPID))) & "] ������"
                        Exit Function
                    End If
                    .SetEntryValue "FSCSPID", lFItemID, lRow              '��λ
                    
'                    lFItemID = GetDCSPID(CNN, Trim(CStr(varDatas(i)(CONST_D01_FDestBinID))), lSPGroupID)
'                    If lFItemID = 0 Then
'                        sErr = "�ջ���λ[" & Trim(CStr(varDatas(i)(CONST_D01_FDestBinID))) & "] ������"
'                        Exit Function
'                    End If
'                    lSPID = lFItemID
'                    .SetEntryValue "FDCSPID", lFItemID, lRow              '��λ
                End If
                
                lFItemID = GetItemID(cnn, Trim(CStr(varDatas(i)(CONST_B01_FCustomerCode))), 5)
                If lFItemID = 0 Then
                    sErr = "������[" & sSourceBillNo & "] :�ջ��ֿ�[" & Trim(CStr(varDatas(i)(CONST_B01_FCustomerCode))) & "] ������"
                    Exit Function
                End If
                .SetEntryValue "FDCStockID", lFItemID, lRow            '�ջ��ֿ�
                lDCStockID = lFItemID
                
                
                lFItemID = GetItemID(cnn, Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))), 4, lUnitID, dRate)
                If lFItemID = 0 Then
                    sErr = "������[" & sSourceBillNo & "] :��Ʒ[" & Trim(CStr(varDatas(i)(CONST_B01_FItemNumber))) & "]������"
                    Exit Function
                End If
                .SetEntryValue "FItemID", lFItemID, lRow     '����
                .SetEntryValue "FUnitID", lUnitID, lRow      '��λ
                
                
                '���źͱ�����
                If Len(CStr(varDatas(i)(CONST_B01_FLotNumber))) <= 0 Then
                    sErr = "������[" & sSourceBillNo & "] :���Ų���Ϊ��"
                    Exit Function
                End If
                .SetEntryValue "FBatchNo", CStr(Trim(varDatas(i)(CONST_B01_FLotNumber))), lRow              ''�ֿ�
                
                If Len(CStr(varDatas(i)(CONST_B01_FProductionDate))) <= 0 Then
                    sErr = "������[" & sSourceBillNo & "] :�������ڲ���Ϊ��"
                    Exit Function
                End If
                .SetEntryValue "FKFDate", CStr(Trim(varDatas(i)(CONST_B01_FProductionDate))), lRow              ''�ֿ�
                
                If Len(CStr(varDatas(i)(CONST_B01_FShelfLife))) <= 0 Then
                    sErr = "������[" & sSourceBillNo & "] :�����ڲ���Ϊ��"
                    Exit Function
                End If
                .SetEntryValue "FKFPeriod", Val(varDatas(i)(CONST_B01_FShelfLife)), lRow              ''�ֿ�
                
                If Len(CStr(varDatas(i)(CONST_B01_FExpiryDate))) <= 0 Then
                    sErr = "������[" & sSourceBillNo & "] :�������ڲ���Ϊ��"
                    Exit Function
                End If
                .SetEntryValue "FPeriodDate", Trim(CStr(varDatas(i)(CONST_B01_FExpiryDate))), lRow '��Ч����


                If Val(CStr(varDatas(i)(CONST_B01_FQty))) <= 0 Then
                    sErr = "������[" & sSourceBillNo & "] :�ƿ���������Ϊ����"
                    Exit Function
                End If
                .SetEntryValue "Fauxqty", Val(varDatas(i)(CONST_B01_FQty)), lRow           '����
                
                
                StrSqlTemp = StrSqlTemp & vbCrLf & "insert T_t_Sales  values('������','" & sSourceBillNo & _
                                    "'," & lDCStockID & ",0,'" & Trim(CStr(varDatas(i)(CONST_B01_FBoxBarcode))) & _
                                    "',1,'" & Format(varDatas(i)(CONST_B01_FDATE), "yyyy-mm-dd") & _
                                    "','" & Format(varDatas(i)(CONST_B01_FDATE), "yyyymmddhhmmss") & _
                                    "',0,0,'" & strUUID & "',1," & lStockID & "," & lSPID & ")"
                
                lRow = lRow + 1

'                blnCheck = True ''�Ƿ������Ч����
            Next
            
            cnn.Execute StrSqlTemp

            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
'                checkBillData sDsn, lInterID, 41, 0, sErr, 16394 '��˵���
                
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
    sErr = "���ɵ�����ʧ�ܣ�" & sErr

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


