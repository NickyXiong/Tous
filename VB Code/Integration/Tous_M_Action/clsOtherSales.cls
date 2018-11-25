VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOtherSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************
'Ŀ��:������������ⵥ
'������rsNewBill
'���أ���
'Created By Anson 2010-08-23
'********************************************************
Public Function CreateOtherSales(ByVal strDSN As String, strUUID As String, ByRef Msg As StringBuilder) As KFO.Dictionary
'    msg.Append "kkk,"
'    Exit Function
    
    Dim strMsg As String
    Dim i As Long
    Dim j As Long
    Dim lngBillInterID As Long    '���浥�����ɺ��InterID
    Dim strInvoiceID As String
    Dim objSave As Object
    Dim objIcInv As Object
    Dim propIcInv As Object
    Dim objtemp As Object
    Dim dctTempHead As Object
    Dim bTrans As Boolean '�����־
    Dim strDate As String
    Dim strErrorInvoice As String
    Dim rsNewBill  As ADODB.Recordset
    Dim dctReturn As Object
    Dim vctTemp As Object
    Dim dctDetail As Object
    Dim TempSRCInterID As Long
    Dim tempCurrency As String
    Set rsNewBill = New ADODB.Recordset
    Dim StrSql As String
    Dim rs As Recordset
    Dim rsBillsDetail As Recordset
    Dim rsBillsCount As Recordset

    Dim QTY As Long
    
    '********��Ҫ���ص�dictionay
    Set dctReturn = New KFO.Dictionary
    Set vctTemp = New KFO.Vector
    Set Msg = New StringBuilder
    
    '****************************************
    
On Error GoTo Err_Lab

    Set objSave = CreateObject("BillDataAccess.GetData")
    Set rsNewBill = New ADODB.Recordset
    'Get data ��ѯһ�ε����˶����ŷ�Ʊ
    StrSql = "select distinct FBillNo from T_t_Sales where FUUID='" & strUUID & "' and ftype = '��������'"
    Set rsBillsCount = objSave.ExecuteSQL(strDSN, StrSql)
    

    i = 0
    'ѭ����ͷ��ֵ
    Do While i < rsBillsCount.RecordCount
        StrSql = ""
        Set rsBillsDetail = Nothing

        StrSql = "select t6.FEmpID,t1.FType,t1.FBillNo,CONVERT(varchar(10), t1.FDate, 23)FDate,t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,DATEADD(day,-t3.fkfperiod,t2.FDate) FProductionDate,t3.fkfperiod,"
        StrSql = StrSql & vbCrLf & "t2.FModel,t3.FUnitID,"
        StrSql = StrSql & vbCrLf & "COUNT(1) FQTY"
        StrSql = StrSql & vbCrLf & "from T_t_Sales t1 inner join T_t_Package t2 on t1.FBarCode =t2.FBoxBarCode and t1.FBarSgin=1"
        StrSql = StrSql & vbCrLf & "inner join t_ICItem t3 on t2.FProductNumber=t3.FBarcode"
        StrSql = StrSql & vbCrLf & "inner join t_Stock t6 on t1.FDefaultWareHouse=t6.FItemID"
        StrSql = StrSql & vbCrLf & "where t1.fuuid='" & strUUID & "' and t1.FType='��������'  and t6.FNumber not like '%C004%'"
        StrSql = StrSql & vbCrLf & "group by t6.FEmpID,t1.FType,t1.FBillNo,CONVERT(varchar(10), t1.FDate, 23),t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,t2.FDate,t3.fkfperiod,"
        StrSql = StrSql & vbCrLf & "t2.FModel , t3.FUnitID"
        StrSql = StrSql & vbCrLf & "Union"
        StrSql = StrSql & vbCrLf & "select t6.FEmpID,t1.FType,t1.FBillNo,CONVERT(varchar(10), t1.FDate, 23)FDate,t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,DATEADD(day,-t3.fkfperiod,t2.FDate) FProductionDate,t3.fkfperiod,"
        StrSql = StrSql & vbCrLf & "t2.FModel,t3.FUnitID,"
        StrSql = StrSql & vbCrLf & "COUNT(1) FQTY"
        StrSql = StrSql & vbCrLf & "from T_t_Sales t1 inner join T_t_Package t2 on t1.FBarCode =t2.FHeBarCode and t1.FBarSgin=0"
        StrSql = StrSql & vbCrLf & "inner join t_ICItem t3 on t2.FProductNumber=t3.FBarcode"
        StrSql = StrSql & vbCrLf & "inner join t_Stock t6 on t1.FDefaultWareHouse=t6.FItemID"
        StrSql = StrSql & vbCrLf & "where t1.fuuid='" & strUUID & "' and t1.FType='��������'  and t6.FNumber not like '%C004%'"
        StrSql = StrSql & vbCrLf & "group by t6.FEmpID,t1.FType,t1.FBillNo,CONVERT(varchar(10), t1.FDate, 23),t1.FDefaultWareHouse,t1.FDefaultSPID,t3.FItemID,t2.FBatch,t2.FDate,t3.fkfperiod,t2.FModel , t3.FUnitID"
    
        Set rsBillsDetail = objSave.ExecuteSQL(strDSN, StrSql)
        
        
        Set rsNewBill = rsBillsDetail
                
        If rsBillsDetail.RecordCount = 0 Then
            dctReturn("success") = False
            dctReturn("errmsg") = "���ɵ���ʱ��������ʧ�ܣ�����ϵ����Ա"
            Set CreateOtherSales = dctReturn
            Exit Function
        End If

        Set objIcInv = CreateObject("BillDataAccess.clsBillPackage")
        '    objIcInv.GetBillEmptyPackage MMTS.PropsString, 86
        objIcInv.GetBillEmptyPackage strDSN, 29
    


        bTrans = False
'        strDate = rs1.Fields("FInvDate").Value
        strInvoiceID = rsBillsDetail.Fields("FBillNo").Value
        
        Set dctTempHead = New KFO.Dictionary
        
        With objIcInv
                'Header
'            .SetValue "FBillNO", rsBillsDetail.Fields("FBillNo").Value
            .setValue "FDate", Format(rsBillsDetail.Fields("FDate").Value, "yyyy-MM-dd")
            .setValue "FSettleDate", Format(rsBillsDetail.Fields("FDate").Value, "yyyy-MM-dd")  '��������
'            .SetValue "FSupplyID", rsBillsDetail.Fields("FSupplyID").Value
            .setValue "FDeptID", 330  '������ʱĬ��ȡֵ
            
            .setValue "FPOMode", 36680 '�ɹ�ģʽ:  Ĭ��Ϊ��ͨ�ɹ�
            .setValue "FPOStyle", 252  '�ɹ���ʽ: Ĭ��Ϊ�޹�
            .setValue "FROB", 1
            .setValue "FBillerID", "16394"
'            .SetValue "FCheckerID", "16394"
        
            .setValue "FSManagerID", rsBillsDetail.Fields("FEmpID").Value
            .setValue "FFManagerID", rsBillsDetail.Fields("FEmpID").Value

        End With
            

        With objIcInv                                   'Entry
        
            For j = 1 To rsNewBill.RecordCount
        
                If rsNewBill.EOF = True Then
                    GoTo SaveBill
                End If
                If rsNewBill.Fields("FBillNo").Value <> strInvoiceID Then
                    GoTo SaveBill
                End If
                strInvoiceID = rsNewBill.Fields("FBillNo").Value
                
                .AddRow j
                
    
                .setValue "FUnitID", rsNewBill.Fields("FUnitID").Value, False, j
                .setValue "FItemID", rsNewBill.Fields("FItemID").Value, False, j '
                .setValue "FDCStockID", rsNewBill.Fields("FDefaultWareHouse").Value, False, j
                
                If Val(rsNewBill.Fields("FDefaultSPID").Value) <> 0 Then
                    .setValue "FDCSPID", rsNewBill.Fields("FDefaultSPID").Value, False, j
                End If
                
    '            QTY = rsNewBill.Fields("Count").Value * rsNewBill.Fields("FModel").Value '������������*���
                
'                .SetValue "FAuxQtyMust", rsNewBill.Fields("FAuxQty").Value, False, j
'                .SetValue "FQtyMust", rsNewBill.Fields("FAuxQty").Value, False, j
                .setValue "FAuxQty", rsNewBill.Fields("FQTY").Value, False, j
                .setValue "FQty", rsNewBill.Fields("FQTY").Value, False, j
'                .SetValue "Fauxprice", rsNewBill.Fields("Fauxprice").Value, False, j
'                .SetValue "FAmount", rsNewBill.Fields("Fauxprice").Value * rsNewBill.Fields("FQTY").Value, False, j
                
                .setValue "FBatchNo", rsNewBill.Fields("FBatch").Value, False, j
                .setValue "FKFDate", rsNewBill.Fields("FProductionDate").Value, False, j
                .setValue "FKFPeriod", rsNewBill.Fields("FKFPeriod").Value, False, j
    
                .setValue "FSourceBillNo", rsNewBill.Fields("FBillNo").Value, False, j
'                .SetValue "FSourceEntryID", rsNewBill.Fields("FSourceEntryID").Value, False, j
'                .SetValue "FSourceInterId", rsNewBill.Fields("FSourceInterId").Value, False, j
                .setValue "FSourceTranType", 0, False, j
                
'                .SetValue "FOrderInterID", rsNewBill.Fields("FSourceInterId").Value, j           '
'                .SetValue "FOrderEntryID", rsNewBill.Fields("FSourceEntryID").Value, j           '
'                .SetValue "FOrderBillNo", rsNewBill.Fields("FBillNo").Value, j
                   
                strInvoiceID = rsNewBill.Fields("FBillNo").Value
                rsNewBill.MoveNext
            Next j
             
        End With
        
SaveBill:
        If strMsg <> "" Then
            GoTo Err_Lab
        End If
        
'        GetObjectContext.SetAbort
        bTrans = True '���������־
        
        Set propIcInv = objIcInv.BillPackage
        Set objtemp = objSave.SaveBill(strDSN, propIcInv) 'Save bill
        If objtemp("success") Then                             'Check bill
            lngBillInterID = objtemp("BillInterID")  '���صĵ�������
            
'            checkBillData strDSN, lngBillInterID, 29, 0, strMsg, 16394 '��˵���
            
'            ���⹺��ⵥInterID��TranType��д��t_t_sales��
            StrSql = "update t_t_sales set FTransType=29,FIsExported=1,FInterID=" & lngBillInterID & " where FBillNo='" & rsBillsCount.Fields("FBillNo").Value & "' and fuuid='" & strUUID & "'"
            objSave.ExecuteSQL strDSN, StrSql
            Msg.Append "�������ⵥ�� " & rsBillsCount.Fields("FBillNo").Value & "  ���ɳɹ������ݱ��Ϊ�� " & objtemp("ReturnBillNo") & vbCrLf
            rsBillsCount.MoveNext
            i = i + 1
        Else
           strMsg = objtemp("ReturnMsg")   '����ʧ��ʱ���صĳ�����Ϣ
           Msg.Append "�������ⵥ�� " & rsBillsCount.Fields("FBillNo").Value & "  ����ʧ��" & vbCrLf
           GoTo Err_Lab
        End If
        
        
'        GetObjectContext.SetComplete
        bTrans = False '�ύ����󣬹ر������־
        
        GoTo Line_lbl
    
Line_lbl:
        Set objIcInv = Nothing
        Set objSave = Nothing
    Loop
        dctReturn("success") = True
        dctReturn("errmsg") = strMsg

        Set CreateOtherSales = dctReturn
        
        
        Exit Function
        
Err_Lab:
    Msg.Append Err.Description
    Msg.Append vbCrLf
    Msg.Append strMsg
    dctReturn("success") = False
    dctReturn("errmsg") = strMsg
    Set CreateOtherSales = dctReturn
    Set objSave = Nothing
End Function


