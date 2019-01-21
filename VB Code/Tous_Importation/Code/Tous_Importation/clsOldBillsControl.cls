VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOldBillsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is Industry Bill Interface Class, made by K3BOSPLUGINSWIZAED"
 
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

Private m_lFCheckerID As Long

Private m_lFDate As Long

Private m_lFShippedDate As Long
Private m_lFDepartureDate As Long
Private m_lFPIWDate As Long


Private m_lSalesDistribution As Long


'Added by Nicky - 20140731
Private m_lFCountryOfEx As Long
Private m_lFContainerNumberEx As Long
Private m_lFSSCCNumberEx As Long
Private m_lFPickTicketNumEx As Long
Private m_lFWeightEx As Long
Private m_lFWeightUOMEx As Long
Private m_lFCubeEx As Long
Private m_lFCubeUOMEx As Long
'Added by Nicky - 20140731

'Added by Nicky - 20140805
Private m_lASNNo As Long
Private m_lStoreNumber As Long
Private m_lCasesShipped As Long
Private m_lTotalWeight As Long
Private m_lTotalPrice As Long
Private m_lReceivingWarehouse As Long
Private m_lArrivalStore As Long
Private m_lPOType As Long
Private m_lShippingMode As Long
'Added by Nicky - 20140805

'Added by Nicky - 20141115
Private m_lSupplier As Long        '��Ӧ��
'Added by Nicky - 20141115

Public Sub Show(ByVal oBillTransfer As Object)

'    MMTS.CheckMts 1
 
    '�ӿ�ʵ��
    'ע��: �˷����������, �����޸�
    Set m_BillTransfer = oBillTransfer
 
End Sub



Private Sub Class_Terminate()

'    If m_lTranType = 71 Or m_lTranType = 81 Then
        m_BillTransfer.BillForm.tlbTool.Buttons.Remove 4
'        m_BillTransfer.BillForm.tlbTool.Buttons.Remove 5
'    End If
 
    '�ͷŽӿڶ���
    'ע��: �˷����������, �����޸�
    Set m_BillTransfer = Nothing

End Sub



Private Sub m_BillTransfer_BillInitialize()

    m_lTranType = m_BillTransfer.SaveVect.Item(1).Value("FTransType")
    
    GetCtlOrdIdx "FCustID", True, m_lCustomerID
    GetCtlOrdIdx "FCurrencyID", True, m_lCurrency
    
    GetCtlOrdIdx "FItemID", False, m_lCtlItemID
    GetCtlOrdIdx "FAuxQty", False, m_lAuxQty
    GetCtlOrdIdx "FCheckerID", True, m_lFCheckerID
    
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
    
    
    GetCtlOrdIdx "FDate", True, m_lFDate

    'Added by Nicky - 20141115
    GetCtlOrdIdx "FSupplyID", True, m_lSupplier        '��Ӧ��
    'Added by Nicky - 20141115

    If m_lTranType = 71 Then
    
        m_BillTransfer.BillForm.tlbTool.Buttons.Add 4, "mnuImportExcel", "Upload", , 18
        m_BillTransfer.BillForm.tlbTool.Buttons(4).Description = "Excel"
        m_BillTransfer.BillForm.tlbTool.Buttons(4).ToolTipText = "Excel"
        
'        m_BillTransfer.BillForm.tlbTool.Buttons.Add 5, "mnuUpdatePrice", "Update", , 20
'        m_BillTransfer.BillForm.tlbTool.Buttons(5).Description = "Update Price"
'        m_BillTransfer.BillForm.tlbTool.Buttons(5).ToolTipText = "Update Price"
    End If
    
End Sub

Private Sub m_BillTransfer_RetEvents(ByVal Para As KFO.IDictionary)
Dim strSQL As String
Dim Res
Dim dict As KFO.Dictionary
Dim strFilter As String
Dim Row As Long
Dim fs As New FileSystemObject
Dim strRemark As String

Dim strFPGO As String

Dim lFossilID As Long

    If Not Para.Lookup("EventID") Then Exit Sub

    If UCase(Para("EventID")) = UCase("Bill_ButtonClick") Then
        If Para("EventIDProp") = 1 Then
            Set dict = Para("Para")
'            m_InterID = m_BillTransfer.SaveVect.Item(1).Value("finterid")

            Select Case UCase(dict("Button").Key)
            '�����ʷ��¼��ѯ
                Case UCase("mnuImportExcel")
                    If m_BillTransfer.GetHeadText(m_lFCheckerID) = "" Then
                        
                        Set frm = New frmOpenFile1
                        frm.Show vbModal
                        If fs.FileExists(strMappingFileName) = True Then
                                                                                
                            ImportMapping strMappingFileName
                            
                        Else
                            If Len(strMappingFileName) > 0 Then
                                MsgBox "File dosen't exist.", vbCritical, "Kingdee Prompt"
                            End If
                        End If
                    End If
                    
                    Row = m_BillTransfer.Grid.ActiveRow
                                     
            End Select

        End If
    End If
End Sub

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

Private Sub ClearFieldData()
Dim i As Integer
On Error GoTo Err
    '�����ͷ�ֶ�
    m_BillTransfer.SetHead m_lContactID, ""
    m_BillTransfer.SetHead m_lPaymentTerm, ""

    '��������ֶ�
    For i = LBound(m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID) To UBound(m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID)
        m_BillTransfer.SetGridText i, m_lCusProductID, ""
        m_BillTransfer.SetGridText i, m_lCtlItemID, ""
    Next
Exit Sub
Err:
    Set rs = Nothing
    
End Sub



'ȡģ����ֶ�FieldName
Private Function GetKeyField(strFieldName As String, bIsHead As Boolean) As String
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim strTranType As String
Dim i As Long
On Error GoTo Err

    strTranType = m_BillTransfer.SaveVect.Item(1).Value("FTransType")
    If bIsHead = True Then
        strSQL = "select t2.FFieldName from ICTransactionType t1 inner join ICTemplate t2 on t1.FTemplateID =t2.FID"
        strSQL = strSQL & vbCrLf & "Where t1.FID = " & strTranType & " and t2.FCaption_CHT ='" & strFieldName & "'"
    Else
        strSQL = "select t2.FFieldName from ICTransactionType t1 inner join ICTemplateEntry t2 on t1.FTemplateID =t2.FID"
        strSQL = strSQL & vbCrLf & "Where t1.FID = " & strTranType & " and t2.FHeadCaption_CHT ='" & strFieldName & "'"
    End If
    Set rs = modPub.ExecuteSQL(strSQL)
    
    If rs.RecordCount > 0 Then
        GetKeyField = rs.Fields("FFieldName").Value
    End If

    Set rs = Nothing
    Exit Function
Err:
    Set rs = Nothing
    MsgBox "GetKeyField error:" & Err.Description, vbInformation, "Kingdee Prompt"
End Function



'Import data from excel file
'strFileName is as Excel file name
Sub ImportMapping(strFileName)
Dim strSQL As String
Dim rsAccount As ADODB.Recordset
Dim strErr As String
Dim i, Row As Long

Dim strPrice As String

Dim xlApp As Excel.Application '����EXCEL��
Dim xlBook As Excel.Workbook '���幤������
Dim xlsheet As Excel.Worksheet '���幤������

On Error Resume Next
    Set xlApp = CreateObject("Excel.Application") '����EXCEL����
    Set xlBook = xlApp.Workbooks().Open(strFileName)
    Set xlsheet = xlBook.Worksheets(1) '��EXCEL������
    
    xlApp.Visible = False
   
    Dim errStr As New StringBuilder
    Dim vec As New Vector
    Dim dic As KFO.Dictionary
    
    i = 2
    Row = 1
'    m_BillInterface.DeleteEntryData 3
    Do While Len(xlsheet.Cells(i, 1)) > 0
        Set dic = New KFO.Dictionary
        dic("FItemNumber") = Trim(xlsheet.Cells(i, 1))
        dic("FQty") = Trim(xlsheet.Cells(i, 2))
        dic("FRetailPrice") = Trim(xlsheet.Cells(i, 3))
        
        vec.Add dic
        i = i + 1
    Loop
    xlBook.Close False
    xlApp.Quit
    Set xlsheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    Dim sql As String
    'Find material number
    Dim rs As ADODB.Recordset
    For i = 1 To vec.Size
        sql = "Select 1 From t_icitem Where FNumber='" & vec(i)("FItemNumber") & "'"
        Set rs = modPub.ExecuteSQL(sql)
        If rs.RecordCount = 0 Then
            errStr.Append "[Row " & (i + 1) & "]SKU:" & vec(i)("FItemNumber") & " is NOT created in ERP system." & Chr(10)
            vec(i)("FItemNumber") = ""
        End If
        Set rs = Nothing
    Next
    '----------------------------------------
    
    If errStr.Length > 0 Then
        Shell "notepad.exe ", vbNormalFocus
        SendKeys errStr.StringValue
        Exit Sub
    End If
    
    Row = 1
    While m_BillTransfer.GetGridText(Row, m_lCtlItemID) <> ""
        Row = Row + 1
    Wend
    i = 2
    For i = 1 To vec.Size
        If vec(i)("FItemNumber") <> "" Then
            m_BillTransfer.SetGridText Row, m_lCtlItemID, vec(i)("FItemNumber")
            m_BillTransfer.SetGridText Row, m_lAuxQty, vec(i)("FQty")
            m_BillTransfer.SetGridText Row, m_lConsignPrice, vec(i)("FRetailPrice")
            m_BillTransfer.SetGridText Row, m_lTaxPrice, vec(i)("FRetailPrice") '* (1 + Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100)
            m_BillTransfer.SetGridText Row, m_lActTaxPrice, vec(i)("FRetailPrice") '* (1 + Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100)
            m_BillTransfer.SetGridText Row, m_lConsignAmount, CDbl(vec(i)("FRetailPrice") * vec(i)("FQty"))
'            m_BillTransfer.SetGridText Row, m_lDiscountAmount, CDbl(vec(i)("FRetailPrice") * vec(i)("FQty")) * Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100
            m_BillTransfer.SetGridText Row, m_lDiscountRate, 0
            m_BillTransfer.SetGridText Row, m_lDiscountAmount, 0
            m_BillTransfer.SetGridText Row, m_lAllAmount, CDbl(vec(i)("FRetailPrice") * vec(i)("FQty")) '* (1 + Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100)
                    
            m_BillTransfer.BillForm.InsertRow , , True
            Row = Row + 1
        End If
    Next
    
    m_BillTransfer.BillForm.DelRow Row
'    m_BillTransfer.FireAfterDelRow Row, False
    

End Sub
