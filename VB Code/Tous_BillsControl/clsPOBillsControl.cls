VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSOBillsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'����������ӿ�. ������е�����, �Դ�������¼�
Private WithEvents m_BillTransfer   As k3BillTransfer.Bill
Attribute m_BillTransfer.VB_VarHelpID = -1

Private m_lCtlItemID As Long        '����
Private m_lCusProductID As Long     '�ͻ���Ʒ����
Private m_lCustomerID As Long       '�ͻ�
Private m_lCurrency As Long         '�ұ�
Private m_lPaymentTerm As Long      '��������
 
Private m_lTransferNumber As Long   'Transfer Number

Private m_lTranType As Long
Private m_lInterID As Long
Private m_lResolveRemark As Long
Private m_lResolveStatus As Long
Private m_lSourceInterID As Long
Private m_lSourceEntryID As Long

Private m_lAuxQty As Long           'ʵ������
Private m_lConsignPrice As Long     '���۵���
Private m_lTaxPrice As Long         '��˰����
Private m_lActTaxPrice As Long      'ʵ�ʺ�˰����
Private m_lConsignAmount As Long    '���۽��
Private m_lDiscountRate As Long     '˰��
Private m_lDiscountAmount As Long   '˰��
Private m_lAllAmount As Long        '��˰�ϼ�

Private m_lTaxAmount As Long        '���۶���˰��

Private m_DCSPID As Long            '�����λ
Private m_DCStockID  As Long        '����ֿ�
Private m_SCStockID  As Long        '�����ֿ�

Private m_BoxNumber  As Long        '����
Private m_BoxModel  As Long        '���

Private m_LotNo  As Long        '����
Private m_KFDate  As Long        '��������
Private m_KFPeriod  As Long        '������
Private m_SCSPID As Long        '�����λ


Private m_lFCheckerID As Long


Public Sub Show(ByVal oBillTransfer As Object)

'    MMTS.CheckMts 1
 
    '�ӿ�ʵ��
    'ע��: �˷����������, �����޸�
    Set m_BillTransfer = oBillTransfer
 
End Sub

Private Sub Class_Terminate()

'    If m_lTranType = 71 Or m_lTranType = 81 Then
'        m_BillTransfer.BillForm.tlbTool.Buttons.Remove 4
'    End If
 
    '�ͷŽӿڶ���
    'ע��: �˷����������, �����޸�
    Set m_BillTransfer = Nothing

End Sub


Private Sub m_BillTransfer_BillInitialize()

    m_lTranType = m_BillTransfer.SaveVect.Item(1).Value("FTransType")
    m_lInterID = m_BillTransfer.SaveVect.Item(1).Value("FInterID")
    
    GetCtlOrdIdx "FCustID", True, m_lCustomerID
    GetCtlOrdIdx "FCurrencyID", True, m_lCurrency
    
    GetCtlOrdIdx "FTransferNumber", True, m_lTransferNumber
    
    GetCtlOrdIdx "FItemID", False, m_lCtlItemID
    GetCtlOrdIdx "FAuxQty", False, m_lAuxQty
    GetCtlOrdIdx "FCheckerID", True, m_lFCheckerID
    GetCtlOrdIdx "FResolvedRemark", True, m_lResolveRemark
    GetCtlOrdIdx "FResolvedStatus", True, m_lResolveStatus
    
    GetCtlOrdIdx "Fauxprice", False, m_lConsignPrice
    GetCtlOrdIdx "FAuxTaxPrice", False, m_lTaxPrice
    GetCtlOrdIdx "FAuxPriceDiscount", False, m_lActTaxPrice
    GetCtlOrdIdx "FAmount", False, m_lConsignAmount
    GetCtlOrdIdx "FCess", False, m_lDiscountRate            '˰��
    GetCtlOrdIdx "FTaxAmount", False, m_lDiscountAmount     '˰��
    GetCtlOrdIdx "FAllAmount", False, m_lAllAmount          '��˰�ϼ�
    
    GetCtlOrdIdx "FTaxAmt", False, m_lTaxAmount     '˰��
    
    GetCtlOrdIdx "FDCStockID", False, m_DCStockID   '����ֿ�
    GetCtlOrdIdx "FDCSPID", False, m_DCSPID         '�����λ
    
    GetCtlOrdIdx "FSCStockID", False, m_SCStockID   '����ֿ�
    GetCtlOrdIdx "FSCSPID", False, m_SCSPID   '����ֿ�
    GetCtlOrdIdx "FBoxNumber", False, m_BoxNumber
    GetCtlOrdIdx "FBoxModel", False, m_BoxModel
    
    GetCtlOrdIdx "FBatchNo", False, m_LotNo
    GetCtlOrdIdx "FKFDate", False, m_KFDate
    GetCtlOrdIdx "FKFPeriod", False, m_KFPeriod
    
'
'    If m_lTranType = 21 Then
'        m_BillTransfer.BillForm.tlbTool.Buttons.Add 4, "mnuImportSD", "Upload", , 18
'        m_BillTransfer.BillForm.tlbTool.Buttons(4).Description = "Upload Excel"
'        m_BillTransfer.BillForm.tlbTool.Buttons(4).ToolTipText = "Upload Excel"
'    End If
'
'    If m_lTranType = 1 Then
'        m_BillTransfer.BillForm.tlbTool.Buttons.Add 4, "mnuImportPR", "Upload", , 18
'        m_BillTransfer.BillForm.tlbTool.Buttons(4).Description = "Upload Excel"
'        m_BillTransfer.BillForm.tlbTool.Buttons(4).ToolTipText = "Upload Excel"
'    End If
    
End Sub



Private Sub m_BillTransfer_GridChange(ByVal Col As Long, ByVal Row As Long, ByVal Value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
Dim BoxNumber As Double
Dim BoxModel As Double

Dim obj As Object
Dim dicCell As KFO.Dictionary
        
    If m_lTranType = 81 Then
    
        If Col = m_SCStockID Then 'ѡ�񷢻��ֿ�ʱ�Զ�����Ĭ�ϲ�λ
            strSql = "select t2.FDefaultSPID,t3.FNumber from t_Stock t1 inner join t_StockPlaceGroup t2 on t1.FSPGroupID=t2.FSPGroupID "
            strSql = strSql & vbCrLf & "inner join t_StockPlace t3 on t2.FDefaultSPID=t3.FSPID Where t1.FItemID = " & m_BillTransfer.EntryCtl(m_SCStockID).DInterID(Row)
            Set rs = modPub.ExecSql(strSql)
            m_BillTransfer.SetGridText Row, m_SCSPID, rs.Fields("FNumber").Value
            m_BillTransfer.SetGridText Row, m_DCSPID, rs.Fields("FNumber").Value
            m_BillTransfer.SetGridText Row, m_DCSPID, rs.Fields("FNumber").Value
        End If
    End If
End Sub




'ȡģ����ֶ�FieldName
Private Function GetKeyField(strFieldName As String, bIsHead As Boolean) As String
Dim strSql As String
Dim rs As ADODB.Recordset
Dim strTranType As String
Dim i As Long
On Error GoTo Err

    strTranType = m_BillTransfer.SaveVect.Item(1).Value("FTransType")
    If bIsHead = True Then
        strSql = "select t2.FFieldName from ICTransactionType t1 inner join ICTemplate t2 on t1.FTemplateID =t2.FID"
        strSql = strSql & vbCrLf & "Where t1.FID = " & strTranType & " and t2.FCaption_CHT ='" & strFieldName & "'"
    Else
        strSql = "select t2.FFieldName from ICTransactionType t1 inner join ICTemplateEntry t2 on t1.FTemplateID =t2.FID"
        strSql = strSql & vbCrLf & "Where t1.FID = " & strTranType & " and t2.FHeadCaption_CHT ='" & strFieldName & "'"
    End If
    Set rs = modPub.ExecSql(strSql)
    
    If rs.RecordCount > 0 Then
        GetKeyField = rs.Fields("FFieldName").Value
    End If

    Set rs = Nothing
    Exit Function
Err:
    Set rs = Nothing
    MsgBox "GetKeyField error:" & Err.Description, vbInformation, "�����ʾ"
End Function


'��λ����
Private Sub FilterItem()
Dim strSql As String
Dim rs As ADODB.Recordset
Dim strFilterPre As String
Dim strFilterLast As String
Dim strFilter As String
Dim i As Long
On Error GoTo Err

    strSql = "select t2.FName FGroupName from t_Group t1 inner join t_user t2 on t1.FGroupID=t2.FUserID"
    strSql = strSql & vbCrLf & "inner join t_user t3 on t1.FUserID=t3.FUserID"
    strSql = strSql & vbCrLf & "where t3.FName='" & MMTS.UserName & "'"
       
    Set rs = modPub.ExecSql(strSql)
    
    strFilterPre = "FSPID in (0,"
    strFilterLast = ")"
        
    rs.Filter = "FGroupName='ҵ����-������'"
    If rs.RecordCount > 0 Then
        strFilter = strFilter & "1,2,3,8,"
    End If
    rs.Filter = ""
    
    rs.Filter = "FGroupName='ҵ����-ҵ���'"
    If rs.RecordCount > 0 Then
        strFilter = strFilter & "1,2,4,5,"
    End If
    rs.Filter = ""
    
    rs.Filter = "FGroupName='ҵ����-��ɼ��'"
    If rs.RecordCount > 0 Then
        strFilter = strFilter & "1,2,6,7,"
    End If
    rs.Filter = ""
    
    If Len(strFilter) > 0 Then
        strFilter = Left(strFilter, Len(strFilter) - 1)
        
        Dim entryctls As Variant
        entryctls = m_BillTransfer.EntryCtl
        entryctls(m_DCSPID).FilterString = strFilterPre & strFilter & strFilterLast
        m_BillTransfer.EntryCtl = entryctls
        
    End If
    
    Exit Sub
Err:
    Set rs = Nothing
    MsgBox "Set item filter error:" & Err.Description, vbInformation, "Kingdee Prompt"
End Sub

'
'Private Sub m_BillTransfer_RetEvents(ByVal Para As KFO.IDictionary)
'Dim strSql As String
'Dim Res
'Dim dict As KFO.Dictionary
'Dim strFilter As String
'Dim Row As Long
'Dim fs As New FileSystemObject
'Dim strRemark As String
'
'    If Not Para.Lookup("EventID") Then Exit Sub
'
'    If Para.Lookup("EventID") = True Then
'
'        Select Case Para("EventID")
'            Case "Bill_FillBillDataExt" '����ѡ�����¼�
''                Call OnAfterSelBill(Cancel)
''                RaiseEvent AfterSelBill(Para("Para")("FieldName"), Para, Cancel)
'                If ShowType = 0 And m_lTranType = 21 Then
'                    Dim dblQty As Double
'                    Dim i As Integer
'                    Dim strLotNo As String
'                    Dim strKFDate As String
'                    Dim strKFPeriod As String
'
'                    Dim lItemID As Long
'                    Dim lStockID As Long
'                    Dim lSPID As Long
'
'                    With m_BillTransfer
'                        For i = LBound(.EntryCtl(m_lCtlItemID).DInterID) To UBound(.EntryCtl(m_lCtlItemID).DInterID) - 1
'                            lItemID = Val(.EntryCtl(m_lCtlItemID).DInterID(i))
'                            lStockID = Val(.EntryCtl(m_DCStockID).DInterID(i))
'                            lSPID = Val(.EntryCtl(m_DCSPID).DInterID(i))
'
'                            If GetLotInfor(lItemID, lStockID, lSPID, strLotNo, strKFDate, strKFPeriod) = True Then
'                                .SetGridText i, m_LotNo, strLotNo
'                                .SetGridText i, m_KFDate, strKFDate
'                                .SetGridText i, m_KFPeriod, strKFPeriod
'                            Else
'                                MsgBox "�Զ���ȡ������Ϣʧ��", vbInformation, "�����ʾ"
'                            End If
'                        Next i
'                    End With
'                End If
'
'                m_BillTransfer.DoRefreshControl
'        End Select
'
'    End If
'
'    If UCase(Para("EventID")) = UCase("Bill_ButtonClick") Then
'        If Para("EventIDProp") = 1 Then
'            Set dict = Para("Para")
''            m_InterID = m_BillTransfer.SaveVect.Item(1).Value("finterid")
'
'            Select Case UCase(dict("Button").Key)
'
''                Case UCase("mnuImportPR")
''                    If m_BillTransfer.GetHeadText(m_lFCheckerID) = "" Then
''
''                        Set frm = New frmOpenFile1
''                        frm.Show vbModal
''                        If fs.FileExists(strMappingFileName) = True Then
''                            ImportSO strMappingFileName
''                        Else
''                            If Len(strMappingFileName) > 0 Then
''                                MsgBox "File dosen't exist!", vbCritical, "�����ʾ"
''                            End If
''                        End If
''                    End If
''
''                    Row = m_BillTransfer.Grid.ActiveRow
''
'
'            End Select
'
'        End If
'    End If
'End Sub


Private Function GetCtlOrdIdx(ByVal strFieldName As String, ByVal bOnHead As Boolean, ByRef nCtlIndex As Long) As Boolean
Dim i As Long
Dim vCtl As Variant
    vCtl = IIf(bOnHead, m_BillTransfer.HeadCtl, m_BillTransfer.EntryCtl)
    For i = LBound(vCtl) To UBound(vCtl)
        If UCase(Trim(vCtl(i).FieldName)) = UCase(strFieldName) Then
            If bOnHead = False Then
                nCtlIndex = vCtl(i).FCtlOrder
            Else
                nCtlIndex = vCtl(i).FCtlIndex
            End If
            
            GetCtlOrdIdx = True
            Exit Function
        End If
    Next
End Function

'Private Sub ClearFieldData()
'Dim I As Integer
'On Error GoTo Err
'    '�����ͷ�ֶ�
'    m_BillTransfer.SetHead m_lContactID, ""
'    m_BillTransfer.SetHead m_lPaymentTerm, ""
'
'    '��������ֶ�
'    For I = LBound(m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID) To UBound(m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID)
'        m_BillTransfer.SetGridText I, m_lCusProductID, ""
'        m_BillTransfer.SetGridText I, m_lCtlItemID, ""
'    Next
'Exit Sub
'Err:
'    Set rs = Nothing
'
'End Sub
