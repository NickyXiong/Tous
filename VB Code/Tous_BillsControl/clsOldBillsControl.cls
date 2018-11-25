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
Private m_SCSPID As Long        '���������λ
Public m_SEOutSP As Long        '����֪ͨ����λ
Private m_SEOutStock As Long    '����֪ͨ���ֿ�
Private strSEOutStockSP As String


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
Dim rsField As ADODB.Recordset


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
    GetCtlOrdIdx "FStockID", False, m_SEOutStock
    
    Set rsField = modPub.ExecSql("select isnull(FFieldName,'')FFieldName from ICTemplateEntry where FID='S02' and (FHeadCaption = '��   λ')")
    If rsField.EOF = False Then
        strSEOutStockSP = rsField.Fields("FFieldName")
    End If
    
    GetCtlOrdIdx strSEOutStockSP, False, m_SEOutSP
    
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


Private Sub m_BillTransfer_BeforeSave(ByVal bNew As Boolean, ReturnCode As Long)
Dim BoxNumber As Double
Dim BoxModel As Double
Dim Qty As Double
Dim Row As Integer
Dim sDsn As String
Dim sDsn2  As String
Dim strErr As String
Dim strSql As String
Dim rs As ADODB.Recordset
Dim obj As Object
Dim blnSave As Boolean
Dim sBillNo As String


'===========================================
'2017-9-4��ʱȡ���������湦��
'===========================================

    With m_BillTransfer
    
        If m_lTranType = 81 And .GetHeadText(m_lFCheckerID) = "" Then
            For Row = LBound(.EntryCtl(m_lCtlItemID).DInterID) To UBound(.EntryCtl(m_lCtlItemID).DInterID)
                
                If .GetGridText(Row, m_lCtlItemID) <> "" Then
                
                    '����Ƿ�Ϊ�������
                    If .GetGridText(Row, m_BoxNumber) <> "" And .GetGridText(Row, m_BoxModel) <> "" Then
                    
                        Qty = Val(.GetGridText(Row, m_lAuxQty))
                        BoxNumber = Val(.GetGridText(Row, m_BoxNumber))
                        BoxModel = Val(.GetGridText(Row, m_BoxModel))
                        If BoxModel <> 0 Then
                            If Qty Mod BoxModel <> 0 Then
                                MsgBox "��" & CStr(Row) & "�У���������Ϊ���ı���,����������", vbCritical, "�����ʾ"
                                ReturnCode = -1
                                Exit Sub
                            End If
                        Else
                            MsgBox "��" & CStr(Row) & "�У���治��Ϊ0,����������", vbCritical, "�����ʾ"
                            ReturnCode = -1
                            Exit Sub
                        End If
                    End If
                    
                    '������Ƿ��㹻
                    strSql = "select isnull(sum(FQty),0)FQty from ICInventory where FItemID=" & .EntryCtl(m_lCtlItemID).DInterID(Row)
                    strSql = strSql & vbCrLf & " and FStockID=" & .EntryCtl(m_SCStockID).DInterID(Row)
                    
                    strSql = "select  (isnull(a.fqty,0)-isnull(b.FReservedQty,0))FAvailableQty from"
                    strSql = strSql & vbCrLf & "(select FItemID,FStockID,isnull(FStockPlaceID,0)FStockPlaceID,isnull(sum(FQty),0)FQty from ICInventory"
                    strSql = strSql & vbCrLf & "Where FItemID = " & .EntryCtl(m_lCtlItemID).DInterID(Row)
                    strSql = strSql & vbCrLf & "and FStockID=" & .EntryCtl(m_SCStockID).DInterID(Row)

                    If Len(.EntryCtl(m_SCSPID).DNumber(Row)) > 0 Then
                        strSql = strSql & vbCrLf & "and FStockPlaceID=" & .EntryCtl(m_SCSPID).DInterID(Row)
                    End If

                    strSql = strSql & vbCrLf & "group by FItemID,FStockID,FStockPlaceID)a"
                    strSql = strSql & vbCrLf & "left Join"
                    strSql = strSql & vbCrLf & "(select t1.FItemID,FSCStockID,isnull(FSCSPID,0) FSCSPID,sum(Fauxqty-FAuxStockQty)FReservedQty from SEOrderEntry t1"
                    strSql = strSql & vbCrLf & "inner join SEOrder t5 on t1.FInterID=t5.FInterID"
                    strSql = strSql & vbCrLf & "Where t1.FInterID<> " & CStr(m_lInterID) & " and t1.FItemID = " & .EntryCtl(m_lCtlItemID).DInterID(Row)
                    strSql = strSql & vbCrLf & "and FSCStockID=" & .EntryCtl(m_SCStockID).DInterID(Row)

                    If Len(.EntryCtl(m_SCSPID).DNumber(Row)) > 0 Then
                        strSql = strSql & vbCrLf & "and FSCSPID=" & .EntryCtl(m_SCSPID).DInterID(Row)
                    End If

                    strSql = strSql & vbCrLf & " and t5.FCancellation=0"
                    strSql = strSql & vbCrLf & "group by t1.FItemID,FSCStockID,FSCSPID) b"
                    strSql = strSql & vbCrLf & "on a.FItemID=b.FItemID and a.FStockID=b.FSCStockID and a.FStockPlaceID=b.FSCSPID"


                    Set rs = modPub.ExecSql(strSql)

                    If rs.RecordCount > 0 Then
                        If Val(.GetGridText(Row, m_lAuxQty)) > Val(rs.Fields("FAvailableQty").Value) Then
                            MsgBox "��" & CStr(Row) & "�У���治�㣬�������µ������޸�����������ѡ��ֿ⡣", vbCritical, "�����ʾ"
                            ReturnCode = -1
                            Exit Sub
                        End If
                    Else
                        MsgBox "��" & CStr(Row) & "�У���治�㣬�������µ������޸�����������ѡ��ֿ⡣", vbCritical, "�����ʾ"
                        ReturnCode = -1
                        Exit Sub
                    End If
                
                End If
                
            Next
      
        End If
        
        
         '����֪ͨ������ǰ�����
            If m_lTranType = 83 And .GetHeadText(m_lFCheckerID) = "" Then
            
'            Set rsField = modPub.ExecSql("select FFieldName from ICTemplateEntry where FID='S02' and (FHeadCaption = '��   λ')")
            If strSEOutStockSP <> "" Then
'                strSEOutStockSP = rsField.Fields("FFieldName")
            Else
                MsgBox "���ҷ���֪ͨ���ֶ�[��   λ]ʧ��", vbCritical, "�����ʾ"
                ReturnCode = -1
                Exit Sub
            End If
            
            For Row = LBound(.EntryCtl(m_lCtlItemID).DInterID) To UBound(.EntryCtl(m_lCtlItemID).DInterID)

                If .GetGridText(Row, m_lCtlItemID) <> "" Then


                    '������Ƿ��㹻
'                    strSQL = "select isnull(sum(FQty),0)FQty from ICInventory where FItemID=" & .EntryCtl(m_lCtlItemID).DInterID(row)
'                    strSQL = strSQL & vbCrLf & " and FStockID=" & .EntryCtl(m_DCStockID).DInterID(row)

                    strSql = "select  (isnull(a.fqty,0)-isnull(b.FReservedQty,0))FAvailableQty from"
                    strSql = strSql & vbCrLf & "(select FItemID,FStockID,isnull(FStockPlaceID,0)FStockPlaceID,isnull(sum(FQty),0)FQty from ICInventory"
                    strSql = strSql & vbCrLf & "Where FItemID = " & .EntryCtl(m_lCtlItemID).DInterID(Row)
                    strSql = strSql & vbCrLf & "and FStockID=" & .EntryCtl(m_SEOutStock).DInterID(Row)

                    If Len(.EntryCtl(m_SEOutSP).DNumber(Row)) > 0 Then
                        strSql = strSql & vbCrLf & "and FStockPlaceID=" & .EntryCtl(m_SEOutSP).DInterID(Row)
                    End If

                    strSql = strSql & vbCrLf & "group by FItemID,FStockID,FStockPlaceID)a"
                    strSql = strSql & vbCrLf & "left Join"
                    strSql = strSql & vbCrLf & "(select t1.FItemID,t1.FStockID FSCStockID,isnull(t1." & strSEOutStockSP & ",0) FSCSPID,sum(Fauxqty-FAuxStockQty)FReservedQty from SEOutStockEntry t1"
                    strSql = strSql & vbCrLf & "inner join SEOutStock t5 on t1.FInterID=t5.FInterID"
                    strSql = strSql & vbCrLf & "Where t1.FInterID<> " & CStr(m_lInterID) & " and t1.FItemID = " & .EntryCtl(m_lCtlItemID).DInterID(Row)
                    strSql = strSql & vbCrLf & "and t1.FStockID=" & .EntryCtl(m_SEOutStock).DInterID(Row)

                    If Len(.EntryCtl(m_SEOutSP).DNumber(Row)) > 0 Then
                        strSql = strSql & vbCrLf & "and isnull(t1." & strSEOutStockSP & ",0)=" & .EntryCtl(m_SEOutSP).DInterID(Row)
                    End If

                    strSql = strSql & vbCrLf & " and t5.FClosed=0"
                    strSql = strSql & vbCrLf & "group by t1.FItemID,t1.FStockID,isnull(t1." & strSEOutStockSP & ",0)) b"
                    strSql = strSql & vbCrLf & "on a.FItemID=b.FItemID and a.FStockID=b.FSCStockID and a.FStockPlaceID=b.FSCSPID"


                    Set rs = modPub.ExecSql(strSql)

                    If rs.RecordCount > 0 Then
                        If Val(.GetGridText(Row, m_lAuxQty)) > Val(rs.Fields("FAvailableQty").Value) Then
                            MsgBox "��" & CStr(Row) & "�У���治�㣬�������µ������޸�����������ѡ��ֿ⡣", vbCritical, "�����ʾ"
                            ReturnCode = -1
                            Exit Sub
                        End If
                    Else
                        MsgBox "��" & CStr(Row) & "�У���治�㣬�������µ������޸�����������ѡ��ֿ⡣", vbCritical, "�����ʾ"
                        ReturnCode = -1
                        Exit Sub
                    End If

                End If

            Next

            End If


'            '����ͻ����ӹ�˾�����������۶���ʱ�Զ��ڶ�Ӧ�ӹ�˾���Զ����ɲɹ�����
'            If m_lTranType = 81 Then
'
'            '���ݿͻ������е�"���"ƥ���ѯ��Ӧ�ӹ�˾���״���
'            StrSql = "select t2.FDBName,t2.FAcctNumber from t_acsetentry t1 inner join KDAcctDB.dbo.t_ad_kdAccount_gl t2 on t1.FACSet=t2.FAcctID"
'            StrSql = StrSql & vbCrLf & "inner join t_Organization t3 on t2.FAcctNumber=t3.FShortName"
'            StrSql = StrSql & vbCrLf & "where UPPER(t3.FName)='" & UCase(.GetHeadText(m_lCustomerID)) & "'"
'
'            Set rs = modPub.ExecSql(StrSql)
'
'            If rs.EOF = False Then
'
'            If MsgBox("��⵽�ͻ�Ϊ�ӹ�˾ [ " & rs.Fields("FAcctNumber").Value & " ], �Ƿ���Ҫ���˶�����ת���ӹ�˾��?", vbYesNo, "Kingdee Prompt") = vbYes Then
'                sDsn = TransfersDsn(rs.Fields("FDBName").Value, MMTS.PropsString)
'            End If
'
'            Set obj = CreateObject("ST_New_M_Action.Action_26")
'            blnSave = obj.SynBillP02(MMTS.PropsString, m_sDsn, m_lInterID, sBillNo, strErr)
'
'            If blnSave = False Then
'                MsgBox "�ӹ�˾�ɹ���������ʧ��: " & strsrr, vbOKOnly, "Kingdee Prompt"
'            Else
'                MsgBox "�ӹ�˾�ɹ��������ɳɹ�: " & sBillNo
'            End If
'
'            End If
'
'
'            End If

    
    
    End With

End Sub

'Private Sub m_BillTransfer_EndSave(ByVal BillNo As String)
'Dim sDsn As String
'Dim sDsn2  As String
'Dim strErr As String
'Dim strSql As String
'Dim rs As ADODB.Recordset
'Dim obj As Object
'Dim blnSave As Boolean
'Dim sBillNo As String
''MMTS.CheckMts 1
'        With m_BillTransfer
'
'         '����ͻ����ӹ�˾�����������۶���ʱ�Զ��ڶ�Ӧ�ӹ�˾���Զ����ɲɹ�����
'            If m_lTranType = 81 Then
'
'            '���ݿͻ������е�"���"ƥ���ѯ��Ӧ�ӹ�˾���״���
'            strSql = "select t2.FDBName,t2.FAcctNumber from t_acsetentry t1 inner join KDAcctDB.dbo.t_ad_kdAccount_gl t2 on t1.FACSet=t2.FAcctID"
'            strSql = strSql & vbCrLf & "inner join t_Organization t3 on t2.FAcctNumber=t3.FShortName"
'            strSql = strSql & vbCrLf & "where UPPER(t3.FName)='" & UCase(.GetHeadText(m_lCustomerID)) & "' OR UPPER(t3.FNumber)='" & UCase(.GetHeadText(m_lCustomerID)) & "'"
'
'            Set rs = modPub.ExecSql(strSql)
'
'            If rs.EOF = False Then
'
'            If MsgBox("��⵽�ͻ�Ϊ�ӹ�˾ [ " & rs.Fields("FAcctNumber").Value & " ], �Ƿ���Ҫ���˶�����ת���ӹ�˾��?", vbYesNo, "Kingdee Prompt") = vbYes Then
'                sDsn = TransfersDsn(rs.Fields("FDBName").Value, MMTS.PropsString)
'            Else
'                Exit Sub
'            End If
'
'            Set obj = CreateObject("ST_New_M_Action.Action_26")
'            blnSave = obj.SynBillP02(MMTS.PropsString, sDsn, BillNo, sBillNo, strErr, rs.Fields("FAcctNumber").Value)
'
'            If blnSave = False Then
'                MsgBox "�ӹ�˾�ɹ���������ʧ��: " & strErr & Err.Description, vbOKOnly, "Kingdee Prompt"
'            Else
'                MsgBox "�ӹ�˾�ɹ��������ɳɹ�: " & sBillNo
'            End If
'
'            End If
'
'
'            End If
'
'        End With
'End Sub

Private Sub m_BillTransfer_GridChange(ByVal Col As Long, ByVal Row As Long, ByVal Value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
Dim BoxNumber As Double
Dim BoxModel As Double

Dim obj As Object
Dim dicCell As KFO.Dictionary
        
    If m_lTranType = 81 Then
    
         If Col = m_lCtlItemID Then  'ѡ������ʱ�Զ��������
            strSql = "select FSQLColumnName from t_ItemPropDesc where FItemClassID=4 and FName='���'"
            Set rs = modPub.ExecSql(strSql)
            If rs.EOF = False Then
                strFieldName = rs.Fields("FSQLColumnName")
            End If
            strSql = "select isnull(" & strFieldName & ",0) BoxModel from t_ICItem where FItemID=" & m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID(Row)
            Set rs = modPub.ExecSql(strSql)
            m_BillTransfer.SetGridText Row, m_BoxModel, rs.Fields("BoxModel")

        End If
        
        '����Ĭ�ϲ�λ������Ĭ�ϲ�λ��������λ��
        If Col = m_SCStockID Then
        If Col = m_SCStockID And (m_BillTransfer.EntryCtl(m_SCStockID).DInterID(Row) <> "" And m_BillTransfer.EntryCtl(m_SCStockID).DInterID(Row) <> 0) And m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID(Row) <> 0 Then 'ѡ�񷢻��ֿ�ʱ�Զ�����Ĭ�ϲ�λ
            strSql = "select t2.FDefaultSPID,t3.FNumber from t_Stock t1 inner join t_StockPlaceGroup t2 on t1.FSPGroupID=t2.FSPGroupID "
            strSql = strSql & vbCrLf & "inner join t_StockPlace t3 on t2.FDefaultSPID=t3.FSPID Where t1.FItemID = " & m_BillTransfer.EntryCtl(m_SCStockID).DInterID(Row)
            Set rs = modPub.ExecSql(strSql)
            If rs.Fields("FNumber").Value <> "*" Then
                m_BillTransfer.SetGridText Row, m_SCSPID, CStr(rs.Fields("FNumber").Value)
            Else
                LockCell Row, m_SCSPID, True
            End If
        End If
        End If
        
        If Col = m_DCStockID Then
        If Col = m_DCStockID And (m_BillTransfer.EntryCtl(m_DCStockID).DInterID(Row) <> "" And m_BillTransfer.EntryCtl(m_DCStockID).DInterID(Row) <> 0) And m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID(Row) <> 0 Then 'ѡ�񷢻��ֿ�ʱ�Զ�����Ĭ�ϲ�λ
            strSql = "select t2.FDefaultSPID,t3.FNumber from t_Stock t1 inner join t_StockPlaceGroup t2 on t1.FSPGroupID=t2.FSPGroupID "
            strSql = strSql & vbCrLf & "inner join t_StockPlace t3 on t2.FDefaultSPID=t3.FSPID Where t1.FItemID = " & m_BillTransfer.EntryCtl(m_DCStockID).DInterID(Row)
            Set rs = modPub.ExecSql(strSql)
            If rs.Fields("FNumber").Value <> "*" Then
                m_BillTransfer.SetGridText Row, m_DCSPID, CStr(rs.Fields("FNumber").Value)
            Else
                LockCell Row, m_DCSPID, True
            End If
            
        End If
        End If
    
        If Col = m_BoxNumber Or Col = m_BoxModel Then
        
            Set obj = m_BillTransfer.BillForm
            
            BoxNumber = Val(m_BillTransfer.GetGridText(Row, m_BoxNumber))
            BoxModel = Val(m_BillTransfer.GetGridText(Row, m_BoxModel))
    '        m_BillTransfer.SetGridText Row, m_lAuxQty, BoxNumber * BoxModel
            
            Set dicCell = New KFO.Dictionary
            dicCell.Value("Name") = BoxNumber * BoxModel
            dicCell.Value("Number") = BoxNumber * BoxModel
            dicCell.Value("Value") = BoxNumber * BoxModel
            
            obj.SetBillDataInput dicCell, "Fauxqty", 1, Row
            
            Set dicCell = Nothing
            Set obj = Nothing
    
        End If
    End If
    
    
    If m_lTranType = 71 Then '�ɹ�����
    
        If Col = m_SCStockID Then
        If Col = m_SCStockID And (m_BillTransfer.EntryCtl(m_SCStockID).DInterID(Row) <> "" And m_BillTransfer.EntryCtl(m_SCStockID).DInterID(Row) <> 0) And m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID(Row) <> 0 Then 'ѡ�񷢻��ֿ�ʱ�Զ�����Ĭ�ϲ�λ
            strSql = "select t2.FDefaultSPID,t3.FNumber from t_Stock t1 inner join t_StockPlaceGroup t2 on t1.FSPGroupID=t2.FSPGroupID "
            strSql = strSql & vbCrLf & "inner join t_StockPlace t3 on t2.FDefaultSPID=t3.FSPID Where t1.FItemID = " & m_BillTransfer.EntryCtl(m_SCStockID).DInterID(Row)
            Set rs = modPub.ExecSql(strSql)
            If rs.Fields("FNumber") <> "*" Then
                m_BillTransfer.SetGridText Row, m_SCSPID, CStr(rs.Fields("FNumber").Value)
            Else
                LockCell Row, m_SCSPID, True
            End If
        End If
        End If
        
        If Col = m_DCStockID Then
        If Col = m_DCStockID And (m_BillTransfer.EntryCtl(m_DCStockID).DInterID(Row) <> "" And m_BillTransfer.EntryCtl(m_DCStockID).DInterID(Row) <> 0) And m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID(Row) <> 0 Then 'ѡ�񷢻��ֿ�ʱ�Զ�����Ĭ�ϲ�λ
            strSql = "select t2.FDefaultSPID,t3.FNumber from t_Stock t1 inner join t_StockPlaceGroup t2 on t1.FSPGroupID=t2.FSPGroupID "
            strSql = strSql & vbCrLf & "inner join t_StockPlace t3 on t2.FDefaultSPID=t3.FSPID Where t1.FItemID = " & m_BillTransfer.EntryCtl(m_DCStockID).DInterID(Row)
            Set rs = modPub.ExecSql(strSql)
            If rs.Fields("FNumber") <> "*" Then
                m_BillTransfer.SetGridText Row, m_DCSPID, CStr(rs.Fields("FNumber").Value)
            Else
                LockCell Row, m_DCSPID, True
            End If
            
        End If
        End If
        
    End If
    
    
    If m_lTranType = 72 Then '����֪ͨ��
    
        If Col = m_SCStockID Then
        If Col = m_SCStockID And (m_BillTransfer.EntryCtl(m_SCStockID).DInterID(Row) <> "" And m_BillTransfer.EntryCtl(m_SCStockID).DInterID(Row) <> 0) And m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID(Row) <> 0 Then   'ѡ�񷢻��ֿ�ʱ�Զ�����Ĭ�ϲ�λ
            strSql = "select t2.FDefaultSPID,t3.FNumber from t_Stock t1 inner join t_StockPlaceGroup t2 on t1.FSPGroupID=t2.FSPGroupID "
            strSql = strSql & vbCrLf & "inner join t_StockPlace t3 on t2.FDefaultSPID=t3.FSPID Where t1.FItemID = " & m_BillTransfer.EntryCtl(m_SCStockID).DInterID(Row)
            Set rs = modPub.ExecSql(strSql)
            If rs.Fields("FNumber") <> "*" Then
                m_BillTransfer.SetGridText Row, m_SCSPID, rs.Fields("FNumber")
            End If
        End If
        End If
        
        If Col = m_SEOutStock Then
        If Col = m_SEOutStock And (m_BillTransfer.EntryCtl(m_SEOutStock).DInterID(Row) <> "" And m_BillTransfer.EntryCtl(m_SEOutStock).DInterID(Row) <> 0) And m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID(Row) <> 0 Then   'ѡ�񷢻��ֿ�ʱ�Զ�����Ĭ�ϲ�λ
            strSql = "select t2.FDefaultSPID,t3.FNumber from t_Stock t1 inner join t_StockPlaceGroup t2 on t1.FSPGroupID=t2.FSPGroupID "
            strSql = strSql & vbCrLf & "inner join t_StockPlace t3 on t2.FDefaultSPID=t3.FSPID Where t1.FItemID = " & m_BillTransfer.EntryCtl(m_SEOutStock).DInterID(Row)
            Set rs = modPub.ExecSql(strSql)
            If rs.Fields("FNumber") <> "*" Then
                m_BillTransfer.SetGridText Row, m_DCSPID, rs.Fields("FNumber")
            Else
                LockCell Row, m_DCSPID, True
            End If
        End If
        End If
        
    End If
    
    If m_lTranType = 83 Then '����֪ͨ��
    
        If Col = m_DCStockID Then
        If Col = m_DCStockID And (m_BillTransfer.EntryCtl(m_DCStockID).DInterID(Row) <> "" And m_BillTransfer.EntryCtl(m_DCStockID).DInterID(Row) <> 0) And m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID(Row) <> 0 Then  'ѡ�񷢻��ֿ�ʱ�Զ�����Ĭ�ϲ�λ
            strSql = "select t2.FDefaultSPID,t3.FNumber from t_Stock t1 inner join t_StockPlaceGroup t2 on t1.FSPGroupID=t2.FSPGroupID "
            strSql = strSql & vbCrLf & "inner join t_StockPlace t3 on t2.FDefaultSPID=t3.FSPID Where t1.FItemID = " & m_BillTransfer.EntryCtl(m_DCStockID).DInterID(Row)
            Set rs = modPub.ExecSql(strSql)
            If rs.Fields("FNumber") <> "*" Then
                m_BillTransfer.SetGridText Row, m_DCSPID, CStr(rs.Fields("FNumber"))
            Else
                LockCell Row, m_DCSPID, True
            End If
        End If
        End If
        
        If Col = m_SEOutStock Then
        If Col = m_SEOutStock And (m_BillTransfer.EntryCtl(m_SEOutStock).DInterID(Row) <> "" And m_BillTransfer.EntryCtl(m_SEOutStock).DInterID(Row) <> 0) _
        And m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID(Row) <> 0 Then 'ѡ�񷢻��ֿ�ʱ�Զ�����Ĭ�ϲ�λ
            strSql = "select t2.FDefaultSPID,t3.FNumber from t_Stock t1 inner join t_StockPlaceGroup t2 on t1.FSPGroupID=t2.FSPGroupID "
            strSql = strSql & vbCrLf & "inner join t_StockPlace t3 on t2.FDefaultSPID=t3.FSPID Where t1.FItemID = " & m_BillTransfer.EntryCtl(m_SEOutStock).DInterID(Row)
            Set rs = modPub.ExecSql(strSql)
            If rs.Fields("FNumber") <> "*" Then
                m_BillTransfer.SetGridText Row, m_SEOutSP, CStr(rs.Fields("FNumber").Value)
            Else
                LockCell Row, m_SEOutSP, True
            End If
        End If
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

Private Sub m_BillTransfer_LoadBillEnd(ByVal ShowType As Long)
'MMTS.CheckMts 1
    If ShowType = 0 And m_lTranType = 21 Then
        Dim dblQty As Double
        Dim i As Integer
        Dim strLotNo As String
        Dim strKFDate As String
        Dim strKFPeriod As String
        
        Dim lItemID As Long
        Dim lStockID As Long
        Dim lSPID As Long
                
        With m_BillTransfer
            For i = LBound(.EntryCtl(m_lCtlItemID).DInterID) To UBound(.EntryCtl(m_lCtlItemID).DInterID) - 1
                lItemID = Val(.EntryCtl(m_lCtlItemID).DInterID(i))
                lStockID = Val(.EntryCtl(m_DCStockID).DInterID(i))
                lSPID = Val(.EntryCtl(m_DCSPID).DInterID(i))
                                
                If GetLotInfor(lItemID, lStockID, lSPID, strLotNo, strKFDate, strKFPeriod) = True Then
                    .SetGridText i, m_LotNo, strLotNo
                    .SetGridText i, m_KFDate, strKFDate
                    .SetGridText i, m_KFPeriod, strKFPeriod
                Else
                    MsgBox "�Զ���ȡ������Ϣʧ��", vbInformation, "�����ʾ"
                End If
            Next i
        End With
    End If
    
    If m_lTranType = 21 Or m_lTranType = 81 Then
'        Call FilterItem
    End If
End Sub

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


Private Sub m_BillTransfer_RetEvents(ByVal Para As KFO.IDictionary)
Dim strSql As String
Dim Res
Dim dict As KFO.Dictionary
Dim strFilter As String
Dim Row As Long
Dim fs As New FileSystemObject
Dim strRemark As String

    If Not Para.Lookup("EventID") Then Exit Sub
    
    If Para.Lookup("EventID") = True Then

        Select Case Para("EventID")
            Case "Bill_FillBillDataExt" '����ѡ�����¼�
'                Call OnAfterSelBill(Cancel)
'                RaiseEvent AfterSelBill(Para("Para")("FieldName"), Para, Cancel)
                If ShowType = 0 And m_lTranType = 21 Then
                    Dim dblQty As Double
                    Dim i As Integer
                    Dim strLotNo As String
                    Dim strKFDate As String
                    Dim strKFPeriod As String
                    
                    Dim lItemID As Long
                    Dim lStockID As Long
                    Dim lSPID As Long
                    
                    With m_BillTransfer
                        For i = LBound(.EntryCtl(m_lCtlItemID).DInterID) To UBound(.EntryCtl(m_lCtlItemID).DInterID) - 1
                            lItemID = Val(.EntryCtl(m_lCtlItemID).DInterID(i))
                            lStockID = Val(.EntryCtl(m_DCStockID).DInterID(i))
                            lSPID = Val(.EntryCtl(m_DCSPID).DInterID(i))
                                            
                            If GetLotInfor(lItemID, lStockID, lSPID, strLotNo, strKFDate, strKFPeriod) = True Then
                                .SetGridText i, m_LotNo, strLotNo
                                .SetGridText i, m_KFDate, strKFDate
                                .SetGridText i, m_KFPeriod, strKFPeriod
                            Else
                                MsgBox "�Զ���ȡ������Ϣʧ��", vbInformation, "�����ʾ"
                            End If
                        Next i
                    End With
                End If
                
                m_BillTransfer.DoRefreshControl
        End Select
            
    End If

    If UCase(Para("EventID")) = UCase("Bill_ButtonClick") Then
        If Para("EventIDProp") = 1 Then
            Set dict = Para("Para")
'            m_InterID = m_BillTransfer.SaveVect.Item(1).Value("finterid")

            Select Case UCase(dict("Button").Key)
                    
'                Case UCase("mnuImportPR")
'                    If m_BillTransfer.GetHeadText(m_lFCheckerID) = "" Then
'
'                        Set frm = New frmOpenFile1
'                        frm.Show vbModal
'                        If fs.FileExists(strMappingFileName) = True Then
'                            ImportSO strMappingFileName
'                        Else
'                            If Len(strMappingFileName) > 0 Then
'                                MsgBox "File dosen't exist!", vbCritical, "�����ʾ"
'                            End If
'                        End If
'                    End If
'
'                    Row = m_BillTransfer.Grid.ActiveRow
'

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

'//��������������Ԫ��
Public Sub LockCell(ByVal Row As Long, Col As Long, lockflag As Boolean)
'    With m_BillTransfer
'        .BillForm.vsEntrys.Col = I2(fieldName)
'        .BillForm.vsEntrys.Row = Row
'        .BillForm.vsEntrys.Lock = lockflag
'    End With
    Dim Entry As Variant
    Set Entry = m_BillTransfer.BillForm.vsEntrys
    Entry.Col = Col
    Entry.Row = Row
    Entry.Lock = lockflag
End Sub

