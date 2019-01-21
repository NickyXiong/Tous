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
 
'定义插件对象接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillTransfer   As k3BillTransfer.Bill
Attribute m_BillTransfer.VB_VarHelpID = -1

Private m_lCtlItemID As Long        '物料
Private m_lCusProductID As Long     '客户产品编码
Private m_lCustomerID As Long       '客户
Private m_lCurrency As Long         '币别
Private m_lPaymentTerm As Long      '付款条件
 
Private m_lTransferNumber As Long   'Transfer Number

Private m_lTranType As Long
Private m_lResolveRemark As Long
Private m_lResolveStatus As Long
Private m_lSourceInterID As Long
Private m_lSourceEntryID As Long

Private m_lAuxQty As Long           '实发数量
Private m_lConsignPrice As Long     '销售单价
Private m_lTaxPrice As Long         '含税单价
Private m_lActTaxPrice As Long      '实际含税单价
Private m_lConsignAmount As Long    '销售金额
Private m_lDiscountRate As Long     '税率
Private m_lDiscountAmount As Long   '税额
Private m_lAllAmount As Long        '价税合计

Private m_lTaxAmount As Long        '销售订单税额

Private m_DCSPID As Long            '调入仓位
Private m_DCStockID  As Long        '调入仓库
Private m_SCStockID  As Long        '调出仓库

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
Private m_lSupplier As Long        '供应商
'Added by Nicky - 20141115

Public Sub Show(ByVal oBillTransfer As Object)

'    MMTS.CheckMts 1
 
    '接口实现
    '注意: 此方法必须存在, 请勿修改
    Set m_BillTransfer = oBillTransfer
 
End Sub



Private Sub Class_Terminate()

'    If m_lTranType = 71 Or m_lTranType = 81 Then
        m_BillTransfer.BillForm.tlbTool.Buttons.Remove 4
        m_BillTransfer.BillForm.tlbTool.Buttons.Remove 5
'    End If
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillTransfer = Nothing

End Sub



Private Sub m_BillTransfer_BillInitialize()

    m_lTranType = m_BillTransfer.SaveVect.Item(1).Value("FTransType")
    
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
    GetCtlOrdIdx "FCess", False, m_lDiscountRate            '税率
    GetCtlOrdIdx "FTaxAmount", False, m_lDiscountAmount     '税额
    GetCtlOrdIdx "FAllAmount", False, m_lAllAmount          '价税合计
    
    GetCtlOrdIdx "FTaxAmt", False, m_lTaxAmount     '税额
    
    GetCtlOrdIdx "FDCStockID", False, m_DCStockID   '调入仓库
    GetCtlOrdIdx "FDCSPID", False, m_DCSPID         '调入仓位
    
    GetCtlOrdIdx "FSCStockID", False, m_SCStockID   '调入仓库
    
    
    GetCtlOrdIdx "FDate", True, m_lFDate
    GetCtlOrdIdx "FShippedDateEx", True, m_lFShippedDate
    GetCtlOrdIdx "FDepartureDate", True, m_lFDepartureDate
    GetCtlOrdIdx "FPIWDate", True, m_lFPIWDate

    GetCtlOrdIdx "FEntrySelfB0300", False, m_lSalesDistribution   'Sales Distribution
    
    'Added by Nicky - 20140731
    GetCtlOrdIdx "FCountryOfEx", False, m_lFCountryOfEx
    GetCtlOrdIdx "FContainerNumberEx", False, m_lFContainerNumberEx
    GetCtlOrdIdx "FSSCCNumberEx", False, m_lFSSCCNumberEx
    GetCtlOrdIdx "FPickTicketNumEx", False, m_lFPickTicketNumEx
    GetCtlOrdIdx "FWeightEx", False, m_lFWeightEx
    GetCtlOrdIdx "FWeightUOMEx", False, m_lFWeightUOMEx
    GetCtlOrdIdx "FCubeEx", False, m_lFCubeEx
    GetCtlOrdIdx "FCubeUOMEx", False, m_lFCubeUOMEx
    'Added by Nicky - 20140731
    

    'Added by Nicky - 20140805
    GetCtlOrdIdx "FASNNo", True, m_lASNNo
    GetCtlOrdIdx "FStoreNumber", True, m_lStoreNumber
    GetCtlOrdIdx "FCasesShippedEx", True, m_lCasesShipped
    GetCtlOrdIdx "FTotalWeightEx", True, m_lTotalWeight
    GetCtlOrdIdx "FTotalPriceEx", True, m_lTotalPrice
    GetCtlOrdIdx "FDestinationID", True, m_lReceivingWarehouse
    GetCtlOrdIdx "FArrivalStore", True, m_lArrivalStore
    GetCtlOrdIdx "FPOType", True, m_lPOType
    GetCtlOrdIdx "FShippingMode", True, m_lShippingMode
    'Added by Nicky - 20140805

    'Added by Nicky - 20141115
    GetCtlOrdIdx "FSupplyID", True, m_lSupplier        '供应商
    'Added by Nicky - 20141115

    If m_lTranType = 71 Then
    
        m_BillTransfer.BillForm.tlbTool.Buttons.Add 4, "mnuImportExcel", "Upload", , 18
        m_BillTransfer.BillForm.tlbTool.Buttons(4).Description = "Upload Excel"
        m_BillTransfer.BillForm.tlbTool.Buttons(4).ToolTipText = "Upload Excel"
        
        m_BillTransfer.BillForm.tlbTool.Buttons.Add 5, "mnuUpdatePrice", "Update", , 20
        m_BillTransfer.BillForm.tlbTool.Buttons(5).Description = "Update Price"
        m_BillTransfer.BillForm.tlbTool.Buttons(5).ToolTipText = "Update Price"
    End If
    
    If m_lTranType = 81 Then
        m_BillTransfer.BillForm.tlbTool.Buttons.Add 4, "mnuImportSO", "Upload", , 18
        m_BillTransfer.BillForm.tlbTool.Buttons(4).Description = "Upload Excel"
        m_BillTransfer.BillForm.tlbTool.Buttons(4).ToolTipText = "Upload Excel"
    End If
    
    If m_lTranType = 10 Or m_lTranType = 29 Then
        m_BillTransfer.BillForm.tlbTool.Buttons.Add 4, "mnuImportStock", "Upload", , 18
        m_BillTransfer.BillForm.tlbTool.Buttons(4).Description = "Upload Excel"
        m_BillTransfer.BillForm.tlbTool.Buttons(4).ToolTipText = "Upload Excel"
    End If
    
    If m_lTranType = 41 Then
        m_BillTransfer.BillForm.tlbTool.Buttons.Add 4, "mnuResolved", "Resolve", , 20
        m_BillTransfer.BillForm.tlbTool.Buttons(4).Description = "Resolve"
        m_BillTransfer.BillForm.tlbTool.Buttons(4).ToolTipText = "Resolve"
    End If
    
    If m_lTranType = 100 Then
        m_BillTransfer.BillForm.tlbTool.Buttons.Add 4, "mnuImportCA", "Upload", , 18
        m_BillTransfer.BillForm.tlbTool.Buttons(4).Description = "Upload Excel"
        m_BillTransfer.BillForm.tlbTool.Buttons(4).ToolTipText = "Upload Excel"
    End If
End Sub

Private Sub m_BillTransfer_BeforeSave(ByVal bNew As Boolean, ReturnCode As Long)
    Dim I As Integer
    Dim strSCNumber As String
    
    Dim FShippedDate As Date
    Dim FDepDate As Date
    Dim FPIWDate As Date
    
    If m_lTranType = 41 Then
        With m_BillTransfer
            If Trim(.GetHeadText(m_lTransferNumber)) = "" Then
                For I = LBound(.EntryCtl(m_lCtlItemID).DInterID) To UBound(.EntryCtl(m_lCtlItemID).DInterID)
                    strSCNumber = .GetGridText(I, m_SCStockID)
                    If InStr(1, strSCNumber, "T-WH") > 0 Then
                        MsgBox "Please enter [Transfer Number(RM)] for RM special adjustment", vbInformation, "Kingdee Prompt"
                        ReturnCode = -1
                    End If
                Next
            End If
            
        End With
    End If
    
    If m_lTranType = 71 Then
        With m_BillTransfer
            If Trim(.GetHeadText(m_lFPIWDate)) = "" Or Trim(.GetHeadText(m_lFDepartureDate)) = "" Or Trim(.GetHeadText(m_lFShippedDate)) = "" Then
                MsgBox "Please enter [PIW Date], [Departure Date] and [Ship Confirm Date]", vbInformation, "Kingdee Prompt"
                ReturnCode = -1
                Exit Sub
            End If
            
            FShippedDate = CDate(.GetHeadText(m_lFShippedDate))
            FDepDate = CDate(.GetHeadText(m_lFDepartureDate))
            FPIWDate = CDate(.GetHeadText(m_lFPIWDate))
                
            If FShippedDate > FDepDate Then
                MsgBox "[Departure Date] should be equal or later than [Ship Confirm Date]", vbInformation, "Kingdee Prompt"
                ReturnCode = -1
                Exit Sub
            ElseIf FDepDate > FPIWDate Then
                MsgBox "[PIW Date] should be equal or later than [Departure Date]", vbInformation, "Kingdee Prompt"
                ReturnCode = -1
                Exit Sub
            ElseIf FShippedDate > FPIWDate Then
                MsgBox "[PIW Date] should be equal or later than [Shipped Confirm Date]", vbInformation, "Kingdee Prompt"
                ReturnCode = -1
                Exit Sub
            End If
            
            'ASN No
            If Trim(.GetHeadText(m_lASNNo)) = "" Then
                MsgBox "[ASN No] should NOT be empty", vbInformation, "Kingdee Prompt"
                ReturnCode = -1
                Exit Sub
            End If
            
            'Store Number
            If Trim(.GetHeadText(m_lStoreNumber)) = "" Then
                MsgBox "[Store Number] should NOT be empty", vbInformation, "Kingdee Prompt"
                ReturnCode = -1
                Exit Sub
            End If
            
            'PO Type
            If Trim(.GetHeadText(m_lPOType)) = "" Then
                MsgBox "[PO Type] should Not be empty", vbInformation, "Kingdee Prompt"
                ReturnCode = -1
                Exit Sub
            End If
            
            'Total Price
            If Val(Trim(.GetHeadText(m_lTotalPrice))) = 0 Then
                MsgBox "[Total Price] can NOT be zero", vbInformation, "Kingdee Prompt"
                ReturnCode = -1
                Exit Sub
            End If
                        
            'Pack to Store or Direct to Store
            If Trim(.GetHeadText(m_lPOType)) <> "Normal PO" Then
                'Arrival Store
                If Trim(.GetHeadText(m_lArrivalStore)) = "" Then
                    MsgBox "[Arrival Store] should NOT be empty", vbInformation, "Kingdee Prompt"
                    ReturnCode = -1
                    Exit Sub
                End If
            End If
            'Pack to Store or Direct to Store
            
            'Added by Nicky - 20140905
            'Direct to Store
            If Trim(.GetHeadText(m_lPOType)) = "Direct to Store PO" Then
                'Arrival Store
                If Trim(.GetHeadText(m_lCurrency)) <> "CNY" Then
                    MsgBox "[Currency] should always be CNY in Direct to Store PO", vbInformation, "Kingdee Prompt"
                    ReturnCode = -1
                    Exit Sub
                End If
            End If
            'Direct to Store
            
            'Normal PO or Pack to Store
            If Trim(.GetHeadText(m_lPOType)) <> "Direct to Store PO" Then
            
                'Shipping Mode
                If Trim(.GetHeadText(m_lShippingMode)) = "" Then
                    MsgBox "[Shipping Mode] should NOT be empty", vbInformation, "Kingdee Prompt"
                    ReturnCode = -1
                    Exit Sub
                End If
                
                'Receiving Warehouse
                If Trim(.GetHeadText(m_lReceivingWarehouse)) = "" Then
                    MsgBox "[Distribution Warehouse] should NOT be empty", vbInformation, "Kingdee Prompt"
                    ReturnCode = -1
                    Exit Sub
                End If
                            
                
                'Cases Shipped
                If Val(Trim(.GetHeadText(m_lCasesShipped))) = 0 Then
                    MsgBox "[Cases Shipped] can NOT be zero", vbInformation, "Kingdee Prompt"
                    ReturnCode = -1
                    Exit Sub
                End If
                                
                'Total Weight
                If Val(Trim(.GetHeadText(m_lTotalWeight))) = 0 Then
                    MsgBox "[Total Weight] can NOT be zero", vbInformation, "Kingdee Prompt"
                    ReturnCode = -1
                    Exit Sub
                End If
                
                For I = LBound(.EntryCtl(m_lCtlItemID).DInterID) To UBound(.EntryCtl(m_lCtlItemID).DInterID)
                    If IsNumeric(.GetGridText(I, m_lFPickTicketNumEx)) = False Then
                        MsgBox "Row " & CStr(I) & ": [Pick Ticket Number] Should be numeric", vbInformation, "Kingdee Prompt"
                        ReturnCode = -1
                    End If
                Next
            End If
            'Normal PO or Pack to Store
            
            For I = LBound(.EntryCtl(m_lCtlItemID).DInterID) To UBound(.EntryCtl(m_lCtlItemID).DInterID)
                If IsNumeric(.GetGridText(I, m_lFSSCCNumberEx)) = False Then
                    MsgBox "Row " & CStr(I) & ": [SSCC Number] Should be numeric", vbInformation, "Kingdee Prompt"
                    ReturnCode = -1
                End If
            Next
            
        End With
    End If
End Sub

Private Sub m_BillTransfer_GridChange(ByVal Col As Long, ByVal Row As Long, ByVal Value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)

    '2014-5-24 Added by Nicky
    '增加核算项目Sales Distribution
    Dim rsTemp As Recordset
    Dim rsDefault As Recordset
    Dim strSQL As String
    
On Error GoTo Err
    If Col = m_lCtlItemID And Len(Value) <> 0 And (m_lTranType = 10 Or m_lTranType = 29) Then
        With m_BillTransfer
                     
            '修改支持Sales Distribution:Division一对多
            strSQL = "select FNumber,FDivCode from t_TB_ItemCategory"
            Set rsTemp = modPub.ExecuteSql(strSQL)
            If rsTemp.RecordCount > 0 Then
                For k = 1 To rsTemp.RecordCount
                    If Trim(rsTemp.Fields("FDivCode").Value) <> "" Then
                        strSQL1 = "select 1 from t_ICItem where FNumber =" & Value
                        strSQL1 = strSQL1 & vbCrLf & " and FDivCode in (" & rsTemp.Fields("FDivCode").Value & ")"
                        Set rsDefault = modPub.ExecuteSql(strSQL1)
                        If rsDefault.RecordCount > 0 Then
'                            .SetEntryValue "FEntrySelfB0300", rsTemp.Fields("FItemID").Value, lRow
                            .SetGridText Row, m_lSalesDistribution, rsTemp.Fields("FNumber").Value
                            Exit For
                        End If
                    End If
                    
                    rsTemp.MoveNext
                Next
    
            End If
            
            
        End With
    End If
    
    Exit Sub
Err:
    
End Sub


Private Sub m_BillTransfer_HeadChange(ByVal CtlIndex As Long, ByVal Value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
    If m_lTranType = 71 And CtlIndex = m_lPOType Then
        If Value = "03" Then
            m_BillTransfer.SetHead m_lCurrency, "CNY"
        End If
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
            '相关历史记录查询
                Case UCase("mnuImportExcel")
                    If m_BillTransfer.GetHeadText(m_lFCheckerID) = "" Then
                        
                        Set frm = New frmOpenFile1
                        frm.Show vbModal
                        If fs.FileExists(strMappingFileName) = True Then
                        
                            If InStr(UCase(m_BillTransfer.GetHeadText(m_lPOType)), "DIRECT") > 0 Then
                                m_BillTransfer.SetHead m_lCurrency, "CNY"
                                lFossilID = 1
                            Else
                                lFossilID = 0
                            End If
                                                        
                            ImportMapping strMappingFileName, lFossilID
                        Else
                            If Len(strMappingFileName) > 0 Then
                                MsgBox "File dosen't exist!", vbCritical, "Kingdee Prompt"
                            End If
                        End If
                    End If
                    
                    Row = m_BillTransfer.Grid.ActiveRow
                    
                Case UCase("mnuUpdatePrice")
                    If m_BillTransfer.GetHeadText(m_lFCheckerID) = "" Then
                        
                        If Trim(m_BillTransfer.GetHeadText(m_lFPIWDate)) = "" Or Trim(m_BillTransfer.GetHeadText(m_lFDepartureDate)) = "" Then
                            MsgBox "Please enter [PIW Date] and [Departure Date] before updating price.", vbInformation, "Kingdee Prompt"
                            Exit Sub
                        End If
                        
                        If MsgBox("Would you like to refresh import price from lastest product master? If so, please click 'Yes' button to close this PO for pricing refreshment." & vbCrLf & vbCrLf _
                                   & "Please verify the price update by re-open the same PO in Kingdee. Thanks", vbYesNo, "Kingdee Prompt") = vbYes Then
                                    
                                                        
                            If InStr(UCase(m_BillTransfer.GetHeadText(m_lPOType)), "IRECT") > 0 Then
                                m_BillTransfer.SetHead m_lCurrency, "CNY"
                                Call m_BillTransfer.BillForm.mnuFileSave_Click
                                
                                strSQL = "update POOrderEntry set Fauxprice=ISNULL(FFPGOPrice,0),FAuxTaxPrice=ISNULL(FFPGOPrice,0)*(1+FCess/100),FAuxPriceDiscount=ISNULL(FFPGOPrice,0)*(1+FCess/100),"
                                strSQL = strSQL & vbCrLf & "FAmount=ISNULL(FFPGOPrice,0)*Fauxqty,FAllAmount=ISNULL(FFPGOPrice,0)*(1+FCess/100)*Fauxqty"
                                strSQL = strSQL & vbCrLf & "from POOrderEntry t1 inner join t_ICItem t2 on t1.FItemID=t2.FItemID"
                                strSQL = strSQL & vbCrLf & "Where t1.FInterID =" & m_BillTransfer.SaveVect.Item(1).Value("FInterID")
    
                                modPub.ExecuteSql strSQL
                                
                                strSQL = "update POOrder set FTotalPriceEx=t2.FTotalPriceEx"
                                strSQL = strSQL & vbCrLf & "from POOrder t1 inner join (select FInterID,SUM(FAllAmount)FTotalPriceEx from POOrderEntry"
                                strSQL = strSQL & vbCrLf & "Where FInterID =" & m_BillTransfer.SaveVect.Item(1).Value("FInterID")
                                strSQL = strSQL & vbCrLf & "group by FInterID) t2 on t1.FInterID=t2.FInterID"
                                strSQL = strSQL & vbCrLf & "Where t1.FInterID =" & m_BillTransfer.SaveVect.Item(1).Value("FInterID")
    
                                modPub.ExecuteSql strSQL
                            Else
                                m_BillTransfer.SetHead m_lCurrency, "USD"
                                Call m_BillTransfer.BillForm.mnuFileSave_Click
                                
                                strSQL = "update POOrderEntry set Fauxprice=FImportPrice,FAuxTaxPrice=FImportPrice,FAuxPriceDiscount=FImportPrice,"
                                strSQL = strSQL & vbCrLf & "FAmount=FImportPrice*Fauxqty,FAllAmount=FImportPrice*Fauxqty"
                                strSQL = strSQL & vbCrLf & "from POOrderEntry t1 inner join t_ICItem t2 on t1.FItemID=t2.FItemID"
                                strSQL = strSQL & vbCrLf & "Where t1.FInterID =" & m_BillTransfer.SaveVect.Item(1).Value("FInterID")
    
                                modPub.ExecuteSql strSQL
                                
                                strSQL = "update POOrder set FTotalPriceEx=t2.FTotalPriceEx"
                                strSQL = strSQL & vbCrLf & "from POOrder t1 inner join (select FInterID,SUM(FAmount)FTotalPriceEx from POOrderEntry"
                                strSQL = strSQL & vbCrLf & "Where FInterID =" & m_BillTransfer.SaveVect.Item(1).Value("FInterID")
                                strSQL = strSQL & vbCrLf & "group by FInterID) t2 on t1.FInterID=t2.FInterID"
                                strSQL = strSQL & vbCrLf & "Where t1.FInterID =" & m_BillTransfer.SaveVect.Item(1).Value("FInterID")
    
                                modPub.ExecuteSql strSQL
                            End If
                                                        
                            Call m_BillTransfer.BillForm.mnuFileExit_Click
                            
                        End If
                    End If
                    
                Case UCase("mnuImportSO")
                    If m_BillTransfer.GetHeadText(m_lFCheckerID) = "" Then
                    
                        Set frm = New frmOpenFile1
                        frm.Show vbModal
                        If fs.FileExists(strMappingFileName) = True Then
                            ImportSO strMappingFileName
                        Else
                            If Len(strMappingFileName) > 0 Then
                                MsgBox "File dosen't exist!", vbCritical, "Kingdee Prompt"
                            End If
                        End If
                    End If
                    
                    Row = m_BillTransfer.Grid.ActiveRow

                Case UCase("mnuImportStock")
                    If m_BillTransfer.GetHeadText(m_lFCheckerID) = "" Then
                    
                        Set frm = New frmOpenFile1
                        frm.Show vbModal
                        If fs.FileExists(strMappingFileName) = True Then
                            ImportOtherStock strMappingFileName
                        Else
                            If Len(strMappingFileName) > 0 Then
                                MsgBox "File dosen't exist!", vbCritical, "Kingdee Prompt"
                            End If
                        End If
                    End If
                    
                    Row = m_BillTransfer.Grid.ActiveRow
                    
                Case UCase("mnuResolved")
                    
                    If m_BillTransfer.GetHeadText(m_lFCheckerID) <> "" Then

                        Set frm = New frmResolved
                        frm.lngInterID = m_BillTransfer.SaveVect.Item(1).Value("FInterID")
                        frm.m_strRemark = m_BillTransfer.GetHeadText(m_lResolveRemark)
                        frm.m_strStatus = m_BillTransfer.GetHeadText(m_lResolveStatus)
                        frm.Show vbModal
                        
                        If frm.m_bUpdated = True Then
                            If m_BillTransfer.GetHeadText(m_lResolveStatus) <> "Y" Then
                                m_BillTransfer.SetHead m_lResolveRemark, frm.m_strRemark
                                m_BillTransfer.SetHead m_lResolveStatus, "Y"
                            Else
                                m_BillTransfer.SetHead m_lResolveRemark, frm.m_strRemark
                                m_BillTransfer.SetHead m_lResolveStatus, "N"
                            End If
                        End If

                    End If

                Case UCase("mnuImportCA")
                    
                    If m_BillTransfer.GetHeadText(m_lFCheckerID) = "" Then

                        If Trim(m_BillTransfer.GetHeadText(m_lFPIWDate)) = "" Or Trim(m_BillTransfer.GetHeadText(m_lFDepartureDate)) = "" Then
                            MsgBox "Please enter [PIW Date] and [Departure Date]", vbInformation, "Kingdee Prompt"
                            Exit Sub
                        End If
            
                        Set frm = New frmOpenFile1
                        frm.Show vbModal
                        If fs.FileExists(strMappingFileName) = True Then
                            ImportCostAdjustment strMappingFileName
                        Else
                            If Len(strMappingFileName) > 0 Then
                                MsgBox "File dosen't exist!", vbCritical, "Kingdee Prompt"
                            End If
                        End If

                    End If
                    
                    Row = m_BillTransfer.Grid.ActiveRow
            End Select

        End If
    End If
End Sub

'过滤客户产品编码
Private Sub FilterItem(ByVal strCustomerNumber As String, ByVal strItemNumber As String)
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim strFilter As String
Dim I As Long
On Error GoTo Err
    
    If Len(strItemNumber) = 0 Then
        strSQL = "select t1.FItemID from t_Item_3009 t1 inner join t_Organization t2 on t1.F_106=t2.FItemID  where t2.FNumber='" & strCustomerNumber & "'"
        Set rs = modPub.ExecuteSql(strSQL)
    Else
        strSQL = "select t1.FItemID from t_Item_3009 t1"
        strSQL = strSQL & vbCrLf & "inner join t_Organization t2 on t1.F_106=t2.FItemID"
        strSQL = strSQL & vbCrLf & "inner join t_ICItem t3 on t1.F_112 = t3.FItemID"
        strSQL = strSQL & vbCrLf & "where t2.FNumber = '" & strCustomerNumber & "' and t3.FNumber ='" & strItemNumber & "'"
        Set rs = modPub.ExecuteSql(strSQL)
    End If
    
    If rs.RecordCount > 0 Then
        strFilter = ""
        strFilter = "t_Item.FItemID in ("
        Do Until rs.EOF
            strFilter = strFilter & "" & rs.Fields("FItemID").Value
            rs.MoveNext
            If rs.EOF = False Then strFilter = strFilter & ","
        Loop
        strFilter = strFilter & ")"
    
        Set rs = Nothing
        
    Else
        strFilter = "t_Item.FItemID in (0)"
    End If
    
    Dim entryctls As Variant
    entryctls = m_BillTransfer.EntryCtl
    entryctls(m_lCusProductID).FilterString = strFilter
    m_BillTransfer.EntryCtl = entryctls
    Exit Sub
Err:
    Set rs = Nothing
    MsgBox "Set item filter error:" & Err.Description, vbInformation, "Kingdee Prompt"
End Sub

Private Function GetCtlOrdIdx(ByVal strFieldName As String, ByVal bOnHead As Boolean, ByRef nCtlIndex As Long) As Boolean
Dim I As Long
Dim vCtl As Variant
    vCtl = IIf(bOnHead, m_BillTransfer.HeadCtl, m_BillTransfer.EntryCtl)
    For I = LBound(vCtl) To UBound(vCtl)
        If UCase(Trim(vCtl(I).FieldName)) = UCase(strFieldName) Then
            If bOnHead = False Then
                nCtlIndex = vCtl(I).FCtlOrder
            Else
                nCtlIndex = vCtl(I).FCtlIndex
            End If
            
            GetCtlOrdIdx = True
            Exit Function
        End If
    Next
End Function

Private Sub ClearFieldData()
Dim I As Integer
On Error GoTo Err
    '清除表头字段
    m_BillTransfer.SetHead m_lContactID, ""
    m_BillTransfer.SetHead m_lPaymentTerm, ""

    '清除表体字段
    For I = LBound(m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID) To UBound(m_BillTransfer.EntryCtl(m_lCtlItemID).DInterID)
        m_BillTransfer.SetGridText I, m_lCusProductID, ""
        m_BillTransfer.SetGridText I, m_lCtlItemID, ""
    Next
Exit Sub
Err:
    Set rs = Nothing
    
End Sub



'取模板表字段FieldName
Private Function GetKeyField(strFieldName As String, bIsHead As Boolean) As String
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim strTranType As String
Dim I As Long
On Error GoTo Err

    strTranType = m_BillTransfer.SaveVect.Item(1).Value("FTransType")
    If bIsHead = True Then
        strSQL = "select t2.FFieldName from ICTransactionType t1 inner join ICTemplate t2 on t1.FTemplateID =t2.FID"
        strSQL = strSQL & vbCrLf & "Where t1.FID = " & strTranType & " and t2.FCaption_CHT ='" & strFieldName & "'"
    Else
        strSQL = "select t2.FFieldName from ICTransactionType t1 inner join ICTemplateEntry t2 on t1.FTemplateID =t2.FID"
        strSQL = strSQL & vbCrLf & "Where t1.FID = " & strTranType & " and t2.FHeadCaption_CHT ='" & strFieldName & "'"
    End If
    Set rs = modPub.ExecuteSql(strSQL)
    
    If rs.RecordCount > 0 Then
        GetKeyField = rs.Fields("FFieldName").Value
    End If

    Set rs = Nothing
    Exit Function
Err:
    Set rs = Nothing
    MsgBox "GetKeyField error:" & Err.Description, vbInformation, "Kingdee Prompt"
End Function

Private Sub m_BillTransfer_LoadBillEnd(ByVal ShowType As Long)
    If m_lTranType = 71 And ShowType = 0 Then
        m_BillTransfer.SetHead m_lFPIWDate, ""
        m_BillTransfer.SetHead m_lFShippedDate, ""
        m_BillTransfer.SetHead m_lFDepartureDate, ""
    End If
End Sub

Private Sub m_BillTransfer_NewBillEnd()
    If m_lTranType = 71 Then
        m_BillTransfer.SetHead m_lFPIWDate, ""
        m_BillTransfer.SetHead m_lFShippedDate, ""
        m_BillTransfer.SetHead m_lFDepartureDate, ""
    End If
End Sub

Private Function GetPRQty(lInterID As Long, lEntryID As Long) As Double
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim strTranType As String
    Dim I As Long
On Error GoTo Err

    strSQL = "select (case when isnull(tt3.FAuxQty,0)-isnull(tt2.FAuxQty,0)<0 then 0 else isnull(tt3.FAuxQty,0)-isnull(tt2.FAuxQty,0) end) FQty"
    '--tt3,取外购入库单数量
    strSQL = strSQL & vbCrLf & "from SEOrderEntry tt1 inner join (select sum(isnull(t3.Fauxqty,0)) FAuxQty,t2.FSourceTranType,t2.FSourceInterId,t2.FSourceEntryID"
    strSQL = strSQL & vbCrLf & "from POOrderEntry t2 left join (select q1.FTranType,q1.FInterID,q2.FEntryID,q2.FSourceTranType,q2.FSourceInterId,q2.FSourceEntryID,q2.FAuxQty"
    strSQL = strSQL & vbCrLf & "from ICStockBill q1 inner join ICStockBillEntry q2 on q1.FInterID=q2.FInterID  where q1.FTranType =1 and q1.FCheckerID>0) t3 on t3.FSourceTranType=71"
    strSQL = strSQL & vbCrLf & "and t3.FSourceInterId=t2.FInterID and t3.FSourceEntryID =t2.FEntryID"
    strSQL = strSQL & vbCrLf & "where t2.FSourceTranType=81 group by t2.FSourceTranType,t2.FSourceInterId,t2.FSourceEntryID) tt3"
    strSQL = strSQL & vbCrLf & "on tt3.FSourceInterId=tt1.FInterID and tt3.FSourceEntryID=tt1.FEntryID"
    '--tt2,取销售出库单数量
    strSQL = strSQL & vbCrLf & "left join (select q2.FSourceTranType,q2.FSourceInterId,q2.FSourceEntryID,sum(isnull(q2.FAuxQty,0)) FAuxQty"
    strSQL = strSQL & vbCrLf & "from ICStockBill q1 inner join ICStockBillEntry q2 on q1.FInterID=q2.FInterID  where q1.FTranType =21 and q1.FCheckerID>0"
    strSQL = strSQL & vbCrLf & "group by q2.FSourceTranType,q2.FSourceInterId,q2.FSourceEntryID)tt2"
    strSQL = strSQL & vbCrLf & "on tt2.FSourceInterId=tt1.FInterID and tt2.FSourceEntryID=tt1.FEntryID"
    strSQL = strSQL & vbCrLf & "Where tt1.FInterID = " & lInterID & " And tt1.FEntryID = " & lEntryID
    Set rs = modPub.ExecuteSql(strSQL)
    
    If rs.RecordCount > 0 Then
        GetPRQty = CDbl(rs.Fields("FQty").Value)
    Else
        GetPRQty = -1
    End If

    Set rs = Nothing
    Exit Function
Err:
    Set rs = Nothing
    MsgBox "GetPRQty error:" & Err.Description, vbInformation, "Kingdee Prompt"
End Function

'采购订单序时簿过滤条件
Public Function GetPOFilter() As String
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim strFilter As String
Dim I, Row As Long
Dim lCol As Long
Dim dct As KFO.Dictionary
On Error GoTo Err
'    m_BillInterface.GetActiveField dct, lCol, Row
    Row = m_BillTransfer.Grid.ActiveRow
    
    strSQL = "select t1.FInterID,t2.FEntryID from POOrder t1 inner join POOrderEntry t2 on t1.FInterID=t2.FInterID "
    strSQL = strSQL & vbCrLf & "inner join t_Organization t3 on t1.FHeadSelfP0252=t3.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_ICItem t4 on t2.FItemID = t4.FItemID"
    strSQL = strSQL & vbCrLf & "inner join t_Item_3009 t5 on t2.FEntrySelfP0265 = t5.FItemID"
    strSQL = strSQL & vbCrLf & "Where t3.FNumber ='" & CStr(m_BillTransfer.GetHeadNumber(m_lCustomerID)) & "'"
    strSQL = strSQL & vbCrLf & " And t4.FNumber ='" & CStr(m_BillTransfer.GetGridText(Row, m_lCtlItemID)) & "'"
    strSQL = strSQL & vbCrLf & " And t5.FName ='" & CStr(m_BillTransfer.GetGridText(Row, m_lCusProductID)) & "'"
    Set rs = modPub.ExecuteSql(strSQL)
        
    If rs.RecordCount > 0 Then
        For I = 1 To rs.RecordCount
            strFilter = strFilter & "(v1.FInterID=" & CStr(rs.Fields("FInterID").Value) & " and u1.FEntryID=" & CStr(rs.Fields("FEntryID").Value) & ")"
        Next
    GetPOFilter = strFilter
    End If
    Exit Function
Err:
    Set rs = Nothing
    MsgBox "Get PO Filter error:" & Err.Description, vbInformation, "Kingdee Prompt"
End Function



'Import data from excel file
'strFileName is as Excel file name
Sub ImportMapping(strFileName, Optional lFossilID = 0)
Dim strSQL As String
Dim rsAccount As ADODB.Recordset
Dim strErr As String
Dim I, Row As Long

Dim strPrice As String

Dim xlApp As Excel.Application '定义EXCEL类
Dim xlBook As Excel.Workbook '定义工件簿类
Dim xlsheet As Excel.Worksheet '定义工作表类

On Error Resume Next
    Set xlApp = CreateObject("Excel.Application") '创建EXCEL对象
    Set xlBook = xlApp.Workbooks().Open(strFileName)
    Set xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表
    
    xlApp.Visible = False
   
    Dim errStr As New StringBuilder
    Dim vec As New Vector
    Dim dic As KFO.Dictionary
    
    If lFossilID = 0 Then
        strPrice = "FImportPrice"
    Else
        strPrice = "FFPGOPrice"
    End If
    
    I = 2
    Row = 1
'    m_BillInterface.DeleteEntryData 3
    Do While Len(xlsheet.Cells(I, 1)) > 0
        Set dic = New KFO.Dictionary
        dic("FStyleNumber") = Trim(xlsheet.Cells(I, 1))
        dic("FColor") = Trim(xlsheet.Cells(I, 2))
        dic("FSize") = Trim(xlsheet.Cells(I, 3))
        dic("FQty") = Trim(xlsheet.Cells(I, 4))
        
        'Added by Nicky - 20140731
        dic("FCountryOfEx") = Trim(xlsheet.Cells(I, 5))
        dic("FContainerNumberEx") = Trim(xlsheet.Cells(I, 6))
        dic("FSSCCNumberEx") = Trim(xlsheet.Cells(I, 7))
        dic("FPickTicketNumEx") = Trim(xlsheet.Cells(I, 8))
        dic("FWeightEx") = Trim(xlsheet.Cells(I, 9))
        dic("FWeightUOMEx") = Trim(xlsheet.Cells(I, 10))
        dic("FCubeEx") = Trim(xlsheet.Cells(I, 11))
        dic("FCubeUOMEx") = Trim(xlsheet.Cells(I, 12))
        'Added by Nicky - 20140731
        
        vec.Add dic
        I = I + 1
    Loop
    xlBook.Close False
    xlApp.Quit
    Set xlsheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    Dim sql As String
    'Find material number
    Dim rs As ADODB.Recordset
    For I = 1 To vec.Size
        sql = "Select t1.FNumber,isnull(t1." & strPrice & ",0) as FPrice From t_icitem t1 Where isnull(t1.FStyleNumber,'')='" _
            & vec(I)("FStyleNumber") & "' and isnull(t1.FColorEx,'')='" & vec(I)("FColor") & "' and isnull(t1.FSizeEx,'')='" & vec(I)("FSize") & "'"
        Set rs = modPub.ExecuteSql(sql)
        If rs.RecordCount = 1 Then
            vec(I)("FItemNumber") = rs.Fields("FNumber").Value
            vec(I)("FRetailPrice") = rs.Fields("FPrice").Value
        ElseIf rs.RecordCount > 1 Then
            errStr.Append "The reocrd in " & (I + 1) & " row:Style Number:" & vec(I)("FStyleNumber") & " and Color:" & vec(I)("FColor") & " and Size:" & vec(I)("FSize") & " record has more than one records" & Chr(10)
            vec(I)("FItemNumber") = ""
        ElseIf rs.RecordCount = 0 Then
            errStr.Append "The reocrd in " & (I + 1) & " row:Style Number:" & vec(I)("FStyleNumber") & " and Color:" & vec(I)("FColor") & " and Size:" & vec(I)("FSize") & " does NOT match to UPC" & Chr(10)
            vec(I)("FItemNumber") = ""
        End If
        Set rs = Nothing
    Next
    '----------------------------------------
    
    Row = 1
    While m_BillTransfer.GetGridText(Row, m_lCtlItemID) <> ""
        Row = Row + 1
    Wend
    I = 2
    For I = 1 To vec.Size
        If vec(I)("FItemNumber") <> "" Then
            m_BillTransfer.SetGridText Row, m_lCtlItemID, vec(I)("FItemNumber")
            m_BillTransfer.SetGridText Row, m_lAuxQty, vec(I)("FQty")
            m_BillTransfer.SetGridText Row, m_lConsignPrice, vec(I)("FRetailPrice")
            m_BillTransfer.SetGridText Row, m_lTaxPrice, vec(I)("FRetailPrice") * (1 + Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100)
            m_BillTransfer.SetGridText Row, m_lActTaxPrice, vec(I)("FRetailPrice") * (1 + Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100)
            m_BillTransfer.SetGridText Row, m_lConsignAmount, CDbl(vec(I)("FRetailPrice") * vec(I)("FQty"))
            m_BillTransfer.SetGridText Row, m_lDiscountAmount, CDbl(vec(I)("FRetailPrice") * vec(I)("FQty")) * Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100
            m_BillTransfer.SetGridText Row, m_lAllAmount, CDbl(vec(I)("FRetailPrice") * vec(I)("FQty")) * (1 + Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100)
            
            'Added by Nicky - 20140731
            m_BillTransfer.SetGridText Row, m_lFCountryOfEx, vec(I)("FCountryOfEx")
            m_BillTransfer.SetGridText Row, m_lFContainerNumberEx, vec(I)("FContainerNumberEx")
            m_BillTransfer.SetGridText Row, m_lFSSCCNumberEx, vec(I)("FSSCCNumberEx")
            m_BillTransfer.SetGridText Row, m_lFPickTicketNumEx, vec(I)("FPickTicketNumEx")
            m_BillTransfer.SetGridText Row, m_lFWeightEx, vec(I)("FWeightEx")
            m_BillTransfer.SetGridText Row, m_lFWeightUOMEx, vec(I)("FWeightUOMEx")
            m_BillTransfer.SetGridText Row, m_lFCubeEx, vec(I)("FCubeEx")
            m_BillTransfer.SetGridText Row, m_lFCubeUOMEx, vec(I)("FCubeUOMEx")
            'Added by Nicky - 20140731
        
            m_BillTransfer.BillForm.InsertRow , , True
            Row = Row + 1
        End If
    Next
    
    m_BillTransfer.BillForm.DelRow Row
'    m_BillTransfer.FireAfterDelRow Row, False
    
    If errStr.Length > 0 Then
        Shell "notepad.exe ", vbNormalFocus
        SendKeys errStr.StringValue
    End If

End Sub

'Import data from excel file
'strFileName is as Excel file name
Sub ImportSO(strFileName)
Dim strSQL As String
Dim rsAccount As ADODB.Recordset
Dim strErr As String
Dim I, Row As Long

Dim xlApp As Excel.Application '定义EXCEL类
Dim xlBook As Excel.Workbook '定义工件簿类
Dim xlsheet As Excel.Worksheet '定义工作表类

On Error Resume Next
    Set xlApp = CreateObject("Excel.Application") '创建EXCEL对象
    Set xlBook = xlApp.Workbooks().Open(strFileName)
    Set xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表
    
    xlApp.Visible = False
    
   
    Dim errStr As New StringBuilder
    Dim vec As New Vector
    Dim dic As KFO.Dictionary
    
    I = 2
    Row = 1
'    m_BillInterface.DeleteEntryData 3
    Do While Len(xlsheet.Cells(I, 5)) > 0
        Set dic = New KFO.Dictionary
        dic("FCustomer") = Trim(xlsheet.Cells(I, 1))
        dic("FCurrency") = Trim(xlsheet.Cells(I, 2))
        dic("FLocation") = Trim(xlsheet.Cells(I, 3))
        dic("FBin") = Trim(xlsheet.Cells(I, 4))
        dic("FUPC") = Trim(xlsheet.Cells(I, 5))
        dic("FQty") = Trim(xlsheet.Cells(I, 6))
        dic("FPrice") = Trim(xlsheet.Cells(I, 7))
        dic("FTotalAmount") = Trim(xlsheet.Cells(I, 8))
        dic("FNetAmount") = Trim(xlsheet.Cells(I, 9))
        dic("FVAT") = Trim(xlsheet.Cells(I, 10))
        vec.Add dic
        I = I + 1
    Loop
    xlBook.Close False
    xlApp.Quit
    Set xlsheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    Dim sql As String
    'Find material number
    Dim rs As ADODB.Recordset
    For I = 1 To vec.Size
        If I = 1 Then
            sql = "select 1 from t_Organization where FNumber='" & vec(I)("FCustomer") & "'"
            Set rs = modPub.ExecuteSql(sql)
            If rs.RecordCount = 0 Then
                errStr.Append "Customer Code:" & vec(I)("FCustomer") & " does NOT exist" & Chr(10)
                vec(I)("FCustomer") = ""
            End If
            Set rs = Nothing
            
            sql = "select 1 from t_Currency where FNumber='" & vec(I)("FCurrency") & "'"
            Set rs = modPub.ExecuteSql(sql)
            If rs.RecordCount = 0 Then
                errStr.Append "Currency Code:" & vec(I)("FCurrency") & " does NOT exist" & Chr(10)
                vec(I)("FCustomer") = ""
            End If
            Set rs = Nothing
            
        End If
        
'        sql = "Select t1.FNumber,t1.FImportPrice From t_icitem t1 Where isnull(t1.FStyleNumber,'')='" _
'            & vec(I)("FStyleNumber") & "' and isnull(t1.FColorEx,'')='" & vec(I)("FColor") & "' and isnull(t1.FSizeEx,'')='" & vec(I)("FSize") & "'"
        sql = "Select FNumber From t_icitem Where FUPCCode='" & vec(I)("FUPC") & "'"
        Set rs = modPub.ExecuteSql(sql)
        If rs.RecordCount = 1 Then
            vec(I)("FItemNumber") = rs.Fields("FNumber").Value
        ElseIf rs.RecordCount > 1 Then
            errStr.Append "The reocrd in " & (I + 1) & " row: UPC Code:" & vec(I)("FUPC") & " record has more than one records" & Chr(10)
            vec(I)("FItemNumber") = ""
        ElseIf rs.RecordCount = 0 Then
            errStr.Append "The reocrd in " & (I + 1) & " row: UPC Code:" & vec(I)("FUPC") & " does NOT match to UPC" & Chr(10)
            vec(I)("FItemNumber") = ""
        End If
        Set rs = Nothing
    Next
    
    '----------------------------------------
    
    If vec(1)("FCustomer") <> "" Then
        m_BillTransfer.SetHead m_lCustomerID, vec(1)("FCustomer")
        m_BillTransfer.SetHead m_lCurrency, vec(1)("FCurrency")
        
        Row = 1
        While m_BillTransfer.GetGridText(Row, m_lCtlItemID) <> ""
            Row = Row + 1
        Wend
        I = 2
        For I = 1 To vec.Size
            If vec(I)("FItemNumber") <> "" Then
                m_BillTransfer.SetGridText Row, m_lCtlItemID, vec(I)("FItemNumber")
                m_BillTransfer.SetGridText Row, m_lAuxQty, vec(I)("FQty")
                m_BillTransfer.SetGridText Row, m_lConsignPrice, vec(I)("FPrice") / 1.17
                m_BillTransfer.SetGridText Row, m_lTaxPrice, vec(I)("FPrice")
                m_BillTransfer.SetGridText Row, m_lActTaxPrice, vec(I)("FPrice")
                m_BillTransfer.SetGridText Row, m_lConsignAmount, CDbl(vec(I)("FNetAmount"))
                m_BillTransfer.SetGridText Row, m_lTaxAmount, CDbl(vec(I)("FVAT"))
'                m_BillTransfer.SetGridText Row, m_lDiscountAmount, CDbl(vec(I)("FRetailPrice") * vec(I)("FQty")) * Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100
                m_BillTransfer.SetGridText Row, m_lAllAmount, CDbl(vec(I)("FTotalAmount"))
                
                m_BillTransfer.SetGridText Row, m_DCStockID, vec(I)("FLocation")
                m_BillTransfer.SetGridText Row, m_DCSPID, vec(I)("FBin")
                
                m_BillTransfer.BillForm.InsertRow , , True
                Row = Row + 1
            End If
        Next
        
        m_BillTransfer.BillForm.DelRow Row
    '    m_BillTransfer.FireAfterDelRow Row, False
    End If
    
    If errStr.Length > 0 Then
        Shell "notepad.exe ", vbNormalFocus
        SendKeys errStr.StringValue
    End If
    

End Sub


'Import data from excel file
'strFileName is as Excel file name
Sub ImportOtherStock(strFileName)
Dim strSQL As String
Dim rsAccount As ADODB.Recordset
Dim strErr As String
Dim I, Row As Long

Dim xlApp As Excel.Application '定义EXCEL类
Dim xlBook As Excel.Workbook '定义工件簿类
Dim xlsheet As Excel.Worksheet '定义工作表类

On Error Resume Next
    Set xlApp = CreateObject("Excel.Application") '创建EXCEL对象
    Set xlBook = xlApp.Workbooks().Open(strFileName)
    Set xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表
    
    xlApp.Visible = False
    
   
    Dim errStr As New StringBuilder
    Dim vec As New Vector
    Dim dic As KFO.Dictionary
    
    '2014-5-24 Added by Nicky
    '增加核算项目Sales Distribution
    Dim rsTemp As Recordset
    Dim rsDefault As Recordset
    
    I = 2
    Row = 1
'    m_BillInterface.DeleteEntryData 3
    Do While Len(xlsheet.Cells(I, 2)) > 0
        Set dic = New KFO.Dictionary
        dic("FUPC") = Trim(xlsheet.Cells(I, 1))
        dic("FStyleNumber") = Trim(xlsheet.Cells(I, 2))
        dic("FColor") = Trim(xlsheet.Cells(I, 3))
        dic("FSize") = Trim(xlsheet.Cells(I, 4))
        dic("FUnitCost") = Trim(xlsheet.Cells(I, 5))
        dic("FLocation") = Trim(xlsheet.Cells(I, 6))
        dic("FBin") = Trim(xlsheet.Cells(I, 7))
        dic("FQty") = Abs(Val(Trim(xlsheet.Cells(I, 8))))
        vec.Add dic
        I = I + 1
    Loop
    xlBook.Close False
    xlApp.Quit
    Set xlsheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    Dim sql As String
    'Find material number
    Dim rs As ADODB.Recordset
    For I = 1 To vec.Size
        sql = "Select t1.FNumber,t1.FImportPrice From t_icitem t1 Where isnull(t1.FStyleNumber,'')='" _
            & vec(I)("FStyleNumber") & "' and isnull(t1.FColorEx,'')='" & vec(I)("FColor") & "' and isnull(t1.FSizeEx,'')='" & vec(I)("FSize") & "'"
        Set rs = modPub.ExecuteSql(sql)
        If rs.RecordCount = 1 Then
            vec(I)("FItemNumber") = rs.Fields("FNumber").Value
'            vec(i)("FRetailPrice") = rs.Fields("FImportPrice").Value
        ElseIf rs.RecordCount > 1 Then
            errStr.Append "The reocrd in " & (I + 1) & " row:Style Number:" & vec(I)("FStyleNumber") & " and Color:" & vec(I)("FColor") & " and Size:" & vec(I)("FSize") & " record has more than one records" & Chr(10)
            vec(I)("FItemNumber") = ""
        ElseIf rs.RecordCount = 0 Then
            errStr.Append "The reocrd in " & (I + 1) & " row:Style Number:" & vec(I)("FStyleNumber") & " and Color:" & vec(I)("FColor") & " and Size:" & vec(I)("FSize") & " does NOT exist!" & Chr(10)
'            errStr.Append "The reocrd in " & (i + 1) & " row:UPC Code:" & vec(i)("FUPC") & " does NOT exist!" & Chr(10)
            vec(I)("FItemNumber") = ""
        End If
        Set rs = Nothing
    Next
    
    '----------------------------------------
    
    Row = 1
    While m_BillTransfer.GetGridText(Row, m_lCtlItemID) <> ""
        Row = Row + 1
    Wend
    I = 2
    For I = 1 To vec.Size
        If vec(I)("FItemNumber") <> "" Then
            m_BillTransfer.SetGridText Row, m_lCtlItemID, vec(I)("FItemNumber")
            m_BillTransfer.SetGridText Row, m_lAuxQty, Val(vec(I)("FQty"))
            m_BillTransfer.SetGridText Row, m_lConsignPrice, Val(vec(I)("FUnitCost"))
'            m_BillTransfer.SetGridText Row, m_lTaxPrice, vec(i)("FRetailPrice") * (1 + Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100)
'            m_BillTransfer.SetGridText Row, m_lActTaxPrice, vec(i)("FRetailPrice") * (1 + Val(m_BillTransfer.GetGridText(Row, m_lDiscountRate)) / 100)
            m_BillTransfer.SetGridText Row, m_lConsignAmount, CDbl(Val(vec(I)("FUnitCost")) * Val(vec(I)("FQty")))
            m_BillTransfer.SetGridText Row, m_DCStockID, vec(I)("FLocation")
            m_BillTransfer.SetGridText Row, m_DCSPID, vec(I)("FBin")
            
            
            '2014-5-24 Added by Nicky
            '增加核算项目Sales Distribution
'            strSQL = "select t2.FNumber  from t_ICItem t1 left join t_TB_ItemCategory t2 "
'            strSQL = strSQL & vbCrLf & "on cast(cast(isnull(t1.FDivCode,0) as int) as nvarchar(5))=t2.FDivCode"
'            strSQL = strSQL & vbCrLf & "where t1.FNumber =" & vec(I)("FItemNumber")
'            Set rsTemp = modPub.ExecuteSql(strSQL)
'            If rsTemp.RecordCount > 0 Then
'                If IsNull(rsTemp.Fields("FNumber").Value) = True Then
'                    Set rsDefault = modPub.ExecuteSql("select FNumber from t_TB_ItemCategory where FName='ADMIN'")
'                    If rsDefault.RecordCount > 0 Then
'                        m_BillTransfer.SetGridText Row, m_lSalesDistribution, rsDefault.Fields("FNumber").Value
'                    End If
'                Else
'                    m_BillTransfer.SetGridText Row, m_lSalesDistribution, rsTemp.Fields("FNumber").Value
'                End If
'            End If
'
                     
            '修改支持Sales Distribution:Division一对多
            bExist = False
            strSQL = "select FNumber,FDivCode from t_TB_ItemCategory"
            Set rsTemp = modPub.ExecuteSql(strSQL)
            If rsTemp.RecordCount > 0 Then
                For k = 1 To rsTemp.RecordCount
                    If Trim(rsTemp.Fields("FDivCode").Value) <> "" Then
                        strSQL1 = "select 1 from t_ICItem where FNumber =" & vec(I)("FItemNumber")
                        strSQL1 = strSQL1 & vbCrLf & " and FDivCode in (" & rsTemp.Fields("FDivCode").Value & ")"
                        Set rsDefault = modPub.ExecuteSql(strSQL1)
                        If rsDefault.RecordCount > 0 Then
'                            .SetEntryValue "FEntrySelfB0300", rsTemp.Fields("FItemID").Value, lRow
                            m_BillTransfer.SetGridText Row, m_lSalesDistribution, rsTemp.Fields("FNumber").Value
                            bExist = True
                            Exit For
                        End If
                    End If
                    
                    rsTemp.MoveNext
                Next
                
                If bExist = False Then
'                            Set rsDefault = CNN.Execute("select FItemID from t_TB_ItemCategory where FName='ADMIN'")
'                            If rsDefault.RecordCount > 0 Then
'                                .SetEntryValue "FEntrySelfB0300", rsDefault.Fields("FItemID").Value, lRow
'                            Else
'                    sErr = "No Sales Distribution information is set for item[" & Trim(CStr(varDatas(I)(CONST_FITEMNUM))) & "]"
'                    Exit Function
'                            End If
                End If
            End If
            
            
            m_BillTransfer.BillForm.InsertRow , , True
            Row = Row + 1
        End If
    Next
    
    m_BillTransfer.BillForm.DelRow Row
'    m_BillTransfer.FireAfterDelRow Row, False
    
    If errStr.Length > 0 Then
        Shell "notepad.exe ", vbNormalFocus
        SendKeys errStr.StringValue
    End If

End Sub


'Import data from excel file
'strFileName is as Excel file name
Sub ImportCostAdjustment(strFileName)
    Dim strSQL As String
    Dim rsAccount As ADODB.Recordset
    Dim strErr As String
    Dim I, Row As Long
    
    Dim xlApp As Excel.Application '定义EXCEL类
    Dim xlBook As Excel.Workbook '定义工件簿类
    Dim xlsheet As Excel.Worksheet '定义工作表类

On Error Resume Next
    Set xlApp = CreateObject("Excel.Application") '创建EXCEL对象
    Set xlBook = xlApp.Workbooks().Open(strFileName)
    Set xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表
    
    xlApp.Visible = False
   
    Dim errStr As New StringBuilder
    Dim vec As New Vector
    Dim dic As KFO.Dictionary
    
    I = 2
    Row = 1
'    m_BillInterface.DeleteEntryData 3
    Do While Len(xlsheet.Cells(I, 1)) > 0
        Set dic = New KFO.Dictionary
        dic("FUPC") = Trim(xlsheet.Cells(I, 1))
        dic("FLocation") = Trim(xlsheet.Cells(I, 2))
        dic("FBin") = Trim(xlsheet.Cells(I, 3))
        dic("FAmount") = Trim(xlsheet.Cells(I, 4))
        vec.Add dic
        I = I + 1
    Loop
    xlBook.Close False
    xlApp.Quit
    Set xlsheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    Dim sql As String
    'Find material number
    Dim rs As ADODB.Recordset
    For I = 1 To vec.Size

'        sql = "Select t1.FNumber,t1.FImportPrice From t_icitem t1 Where isnull(t1.FStyleNumber,'')='" _
'            & vec(I)("FStyleNumber") & "' and isnull(t1.FColorEx,'')='" & vec(I)("FColor") & "' and isnull(t1.FSizeEx,'')='" & vec(I)("FSize") & "'"
        sql = "Select FNumber From t_icitem Where FUPCCode='" & vec(I)("FUPC") & "'"
        Set rs = modPub.ExecuteSql(sql)
        If rs.RecordCount = 1 Then
            vec(I)("FItemNumber") = rs.Fields("FNumber").Value
        ElseIf rs.RecordCount > 1 Then
            errStr.Append "The reocrd in " & (I + 1) & " row: UPC Code:" & vec(I)("FUPC") & " record has more than one records" & Chr(10)
            vec(I)("FItemNumber") = ""
        ElseIf rs.RecordCount = 0 Then
            errStr.Append "The reocrd in " & (I + 1) & " row: UPC Code:" & vec(I)("FUPC") & " does NOT match to UPC" & Chr(10)
            vec(I)("FItemNumber") = ""
        End If
        Set rs = Nothing
    Next
    
    '----------------------------------------
    
    If vec(1)("FUPC") <> "" Then
    
        Row = 1
        While m_BillTransfer.GetGridText(Row, m_lCtlItemID) <> ""
            Row = Row + 1
        Wend
        
        If Row <> 1 Then
            m_BillTransfer.BillForm.InsertRow , , True
        End If
        
        For I = 1 To vec.Size
            If vec(I)("FItemNumber") <> "" Then
                m_BillTransfer.SetGridText Row, m_lCtlItemID, vec(I)("FItemNumber")
                m_BillTransfer.SetGridText Row, m_lConsignAmount, CDbl(vec(I)("FAmount"))
                
                m_BillTransfer.SetGridText Row, m_DCStockID, vec(I)("FLocation")
                m_BillTransfer.SetGridText Row, m_DCSPID, vec(I)("FBin")
                
                m_BillTransfer.BillForm.InsertRow , , True
                Row = Row + 1
            End If
        Next
        
        m_BillTransfer.BillForm.DelRow Row
    '    m_BillTransfer.FireAfterDelRow Row, False
    End If
    
    If errStr.Length > 0 Then
        Shell "notepad.exe ", vbNormalFocus
        SendKeys errStr.StringValue
    End If
    

End Sub




