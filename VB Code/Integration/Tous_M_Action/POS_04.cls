VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************
'Function:������������ⵥ����� (POS TO ERP)
'Date:
'*******************************************************

Option Explicit

Private datasource As CDataSource
Private m_oParse As CParse
Private strLanguage As String


''/* ���۳��ⵥ Sales
'Private Const CONST_B01_FBillNo As Long = 0 '���۵���
'
''Added by Nicky  - 20150504
''�����Ƿ����������ֶ�
'Private Const CONST_B01_FType As Long = 1 '���۳��ⵥ����
''Added by Nicky  - 20150504
'
'Private Const CONST_B01_FDATE As Long = 2 '���۳��ⵥ����
'Private Const CONST_B01_FEntryID  As Long = 3 '����/�˻�
'Private Const CONST_B01_FItemNumber  As Long = 4 '����
'Private Const CONST_B01_FUOM As Long = 5 '���ϴ���
'Private Const CONST_B01_FLotNumber As Long = 6 ' ����
'Private Const CONST_B01_FProductionDate As Long = 7 'ԭ�ۣ���˰��
'Private Const CONST_B01_FShelfLife As Long = 8 'ʵ�����۽���˰��
'Private Const CONST_B01_FExpiryDate As Long = 9 'ԭ���۽���˰��=ԭ�ۣ���˰��*����
'Private Const CONST_B01_FBoxBarcode As Long = 10 ' �ۿ۽�����˰��
'Private Const CONST_B01_FQty As Long = 11 ' �ۿ۽���˰��
'Private Const CONST_B01_FSPID As Long = 12 ' �ۿ���
'Private Const CONST_B01_FReleaseDate As Long = 13 ' ����˰���۽��
'Private Const CONST_B01_FCustomerCode As Long = 14 ' ˰��
'Private Const CONST_B01_FCustomerName As Long = 15 ' ˰��
'Private Const CONST_B01_FAddress As Long = 16 ' ����Ա
'Private Const CONST_B01_FContact As Long = 17 ' ����Ա
'Private Const CONST_B01_FPhone As Long = 18 ' ����Ա


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
' Procedure : SyncStockAdjustment
' DateTime  :
' Author    :
' Purpose   : ������������ⵥ����� (POS TO ERP)
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
    
On Error GoTo SynBillA97B09_Error

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
        
        If Val(varDatas(0)(3)) = 1 Then
            bRet = .NewBill(10)
            lTranType = 10
        ElseIf Val(varDatas(0)(3)) = -1 Then
            bRet = .NewBill(29)
            lTranType = 29
        End If

        If bRet = True Then
            For i = LBound(varDatas) To UBound(varDatas)

                If i = 0 Then '��ͷ
                
                    '���POS Number�Ƿ���ڣ������������ظ�����
                    Set rs = ExecSQL("select 1 from ICStockBill where FTranType=" & lTranType & " and isnull(FPOSNumber,'')='" & CStr(varDatas(i)(1)) & "'", sDsn)
                    If rs.RecordCount > 0 Then
                        sErr = "POS Number[" & CStr(varDatas(i)(1)) & "] has existed in ERP system."
                        GoTo SynBillA97B09_Error
                    End If
                    .SetHeadValue "FPOSNumber", CStr(varDatas(i)(1))  '����
                    
                    .SetHeadValue "FDate", CStr(varDatas(i)(4))  '����
                    
                    
                    '��鲿���Ƿ����
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 2)
                    If lFItemID = 0 Then
                        sErr = "Department Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillA97B09_Error
                    End If
                    .SetHeadValue "FDeptID", lFItemID
                    
                    
                    '���ְԱ�Ƿ����
                    lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 3)
                    If lFItemID = 0 Then
                        sErr = "Employee Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                        GoTo SynBillA97B09_Error
                    End If
                    .SetHeadValue "FEmpID", lFItemID
                    .SetHeadValue "FFManagerID", lFItemID
                    .SetHeadValue "FSManagerID", lFItemID
                    
                    '���Reason Code
                    Set rs = ExecSQL("select FInterID from t_submessage where FTypeID=11400 and fid='" & CStr(varDatas(i)(2)) & "'", sDsn)
                    If rs.RecordCount <= 0 Then
                        sErr = "Reason Code[" & CStr(varDatas(i)(2)) & "] is NOT created in ERP system."
                        GoTo SynBillA97B09_Error
                    End If
                    .SetHeadValue "FAdjReason", rs.Fields("FInterID")
                
                
                End If
                
                
                '���帳ֵ

                .AddRow lRow
                
                ROW = lRow
                
                
                 '��������Ƿ����
                 lFItemID = GetItemIDByNumber(CNN, CStr(varDatas(i)(5)), 4, lUnitID)
                 If lFItemID = 0 Then
                     sErr = "Material Master Data [" & CStr(varDatas(i)(5)) & "] is NOT created in ERP system."
                     GoTo SynBillA97B09_Error
                 End If
                .SetEntryValue "FItemID", lFItemID, lRow           '����
                .SetEntryValue "FUnitID", lUnitID, lRow            '��λ
                
                '�����������0
                If Val(varDatas(i)(6)) <= 0 Then
                    sErr = "The quantity can NOT be zero."
                    GoTo SynBillA97B09_Error
                End If
                .SetEntryValue "Fauxqty", Val(varDatas(i)(6)) * lROB, lRow           '����
                .SetEntryValue "FQty", Val(varDatas(i)(6)) * lROB, lRow          'Ӧ������
                
                '�ֿ�
                lFItemID = GetItemID(CNN, CStr(varDatas(i)(0)), 5)
                If lFItemID = 0 Then
                     sErr = "Warehouse Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                     GoTo SynBillA97B09_Error
                 End If
                .SetEntryValue "FDCStockID", lFItemID, lRow        '�ջ��ֿ�
                
'                 If Val(rsBillInfo.Fields("FIsStockMgr").Value) <> 0 And Val(rsBillInfo.Fields("FBinID").Value) = 0 Then
'                     sErr = "Bin Master Data [" & rsBillInfo.Fields("FBinCode").Value & "] is NOT created in ERP system or does not in the right bin group."
'                     GoTo SynBillA97B09_Error
'                 End If
'                .SetEntryValue "FDCSPID", rsBillInfo.Fields("FBinID").Value, lRow        '�ջ���λ

                .SetEntryValue "FNote", CStr(varDatas(i)(8)), lRow          ''��ע
                                
                lRow = lRow + 1
                
NextRow:
'                blnCheck = True ''�Ƿ������Ч����
            Next
            
            bRet = .SaveBill(lInterID, sBillNo, sErr)
            If bRet = True Then
                checkBillData sDsn, lInterID, lTranType, 0, sErr, 16394 '��˵���
                
            End If
        End If
    End With

    Set clsBill = Nothing
    Set CNN = Nothing
    SyncStockAdjustment = True
    GetObjectContext.SetComplete

   Exit Function

SynBillA97B09_Error:
    SyncStockAdjustment = False
    sErr = "Generate ��" & sErr & Err.Description
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





