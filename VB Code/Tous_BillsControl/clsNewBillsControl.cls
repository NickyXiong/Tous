VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewBillsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is BillEvent Interface Class, made by K3BOSPLUGINSWIZAED"
 
'定义 BillEvent 接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillInterface  As BillEvent
Attribute m_BillInterface.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillInterface As Object)
 
    'BillEvent 接口实现
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = oBillInterface
 
End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = Nothing

End Sub

Private Sub m_BillInterface_BeforeSave(bCancel As Boolean)

''移仓申请单保存前检查库存
'            If m_BillInterface.FID = 200000018 And Val(m_BillInterface.GetFieldValue("FCheckerID", 1)) = "" Then
'            For Row = 1 To m_BillInterface.BillEntrys(2).GridMaxDataRowNum
'
'                If Val(m_BillInterface.GetFieldValue("FItemID", Row)) <> "" Then
'
'
'                    '检查库存是否足够
''                    strSQL = "select isnull(sum(FQty),0)FQty from ICInventory where FItemID=" & .EntryCtl(m_lCtlItemID).DInterID(row)
''                    strSQL = strSQL & vbCrLf & " and FStockID=" & .EntryCtl(m_DCStockID).DInterID(row)
'
'                    strSql = "select  (isnull(a.fqty,0)-isnull(b.FReservedQty,0))FAvailableQty from"
'                    strSql = strSql & vbCrLf & "(select FItemID,FStockID,FStockPlaceID,isnull(sum(FQty),0)FQty from ICInventory"
'                    strSql = strSql & vbCrLf & "Where FItemID = " & .EntryCtl(m_lCtlItemID).DInterID(Row)
'                    strSql = strSql & vbCrLf & "and FStockID=" & Val(m_BillInterface.GetFieldValue("FStockID", Row))
'
'                    If Len(Val(m_BillInterface.GetFieldValue("FSPID", Row))) > 0 Then
'                        strSql = strSql & vbCrLf & "and FStockPlaceID=" & Val(m_BillInterface.GetFieldValue("FSPID", Row))
'
'                    End If
'
'                    strSql = strSql & vbCrLf & "group by FItemID,FStockID,FStockPlaceID)a"
'                    strSql = strSql & vbCrLf & "left Join"
'                    strSql = strSql & vbCrLf & "(select t1.FItemID,t1.FStockID FSCStockID,t1.FEntrySelfS0244 FSCSPID,sum(Fauxqty-FAuxStockQty)FReservedQty from SEOutStockEntry t1"
'                    strSql = strSql & vbCrLf & "inner join SEOutStock t5 on t1.FInterID=t5.FInterID"
'                    strSql = strSql & vbCrLf & "Where t1.FInterID<> " & CStr(m_lInterID) & " and t1.FItemID = " & .EntryCtl(m_lCtlItemID).DInterID(Row)
'                    strSql = strSql & vbCrLf & "and t1.FStockID=" & .EntryCtl(m_SEOutStock).DInterID(Row)
'
'                    If Len(.EntryCtl(m_SEOutSP).DNumber(Row)) > 0 Then
'                        strSql = strSql & vbCrLf & "and FEntrySelfS0244=" & .EntryCtl(m_SEOutSP).DInterID(Row)
'                    End If
'
'                    strSql = strSql & vbCrLf & " and t5.FCancellation=0"
'                    strSql = strSql & vbCrLf & "group by t1.FItemID,t1.FStockID,FEntrySelfS0244) b"
'                    strSql = strSql & vbCrLf & "on a.FItemID=b.FItemID and a.FStockID=b.FSCStockID and a.FStockPlaceID=b.FSCSPID"
'
'
'                    Set rs = modPub.ExecSql(strSql)
'
'                    If rs.RecordCount > 0 Then
'                        If Val(.GetGridText(Row, m_lAuxQty)) > Val(rs.Fields("FAvailableQty").Value) Then
'                            MsgBox "第" & CStr(Row) & "行：库存不足，不允许下单。请修改数量或重新选择仓库。", vbCritical, "金蝶提示"
'                            ReturnCode = -1
'                            Exit Sub
'                        End If
'                    Else
'                        MsgBox "第" & CStr(Row) & "行：库存不足，不允许下单。请修改数量或重新选择仓库。", vbCritical, "金蝶提示"
'                        ReturnCode = -1
'                        Exit Sub
'                    End If
'
'                End If
'
'            Next
'
'            End If

End Sub

Private Sub m_BillInterface_Change(ByVal dct As KFO.IDictionary, ByVal dctFld As KFO.IDictionary, ByVal Col As Long, ByVal Row As Long, Cancel As Boolean)
 
    'TODO: 请在此处添加代码响应事件

Dim CustID As Long  '客户ID
Dim ItemID As Long  '物料ID
Dim BoxSize As Long '箱规
Dim BoxQty As Long  '箱数
Dim strSql As String
Dim CustTypeID As Long  '客户分类ID
Dim Amt As Double   '价税合计
Dim SumAmt As Double    '价税合计总额
Dim StockID As Long     '仓库
Dim Volume As Double    '体积
Dim Weight As Double    '重量
Dim BoxModel As Long    '箱规
Dim Qty As Long     '数量
Dim i As Long   '行数
Dim strFieldName As String
Dim rs As ADODB.Recordset
Dim dctField As KFO.Dictionary
Dim KFDate As Date '生产日期
Dim KFPeriod As Double '保质期
Dim PeriodDate As String '到期日
Dim dblPrice As Double

    If m_BillInterface.FID = 1000019 Then '销售合同

        '获取客户和物料ID
        CustID = Val(m_BillInterface.GetFieldValue("FCustomer", 1))
        ItemID = Val(m_BillInterface.GetFieldValue("FProductID2", Row))
        
        
        If dct("FKey") = "FProductID2" And ItemID <> 0 Then '选择物料时自动带出箱规
            strSql = "select FSQLColumnName from t_ItemPropDesc where FItemClassID=4 and FName='箱规'"
            Set rs = m_BillInterface.K3Lib.GetData(strSql)
            If rs.EOF = False Then
                strFieldName = rs.Fields("FSQLColumnName")
            End If
            strSql = "select isnull(" & strFieldName & ",0) BoxModel from t_ICItem where FItemID=" & ItemID
            Set rs = m_BillInterface.K3Lib.GetData(strSql)
            
            m_BillInterface.SetFieldValue "FBoxModel", rs.Fields("BoxModel"), Row
        End If
        
        
        If dct("FKey") = "FBoxNumber" And CustID <> 0 Then '变更数量时自动带出对应单价,以及计算数量 数量=箱规*箱数
        
            BoxQty = Val(m_BillInterface.GetFieldValue("FBoxNumber", Row))
            BoxSize = Val(m_BillInterface.GetFieldValue("FBoxModel", Row))
        
            Set rs = m_BillInterface.K3Lib.GetData("select FTypeID from  t_Organization where FItemID=" & CustID)
            CustTypeID = rs.Fields("FTypeID")
        
            Set rs = m_BillInterface.K3Lib.GetData("select FPrice from vw_ICPrcPly_CTtoI where FParentID=" & CustTypeID & " and FItemID=" & ItemID & " and (FBegQty<=" & BoxQty & " and FEndQty>=" & BoxQty & ")")
            If rs.EOF Then
                dblPrice = 0
                m_BillInterface.SetFieldValue "FTaxPriceFor", 0, Row
            Else
                '填单价
                dblPrice = rs.Fields("FPrice").Value
                m_BillInterface.SetFieldValue "FTaxPriceFor", rs.Fields("FPrice").Value, Row
            End If
        
            '填数量
            m_BillInterface.SetFieldValue "FQuantity", BoxQty * BoxSize, Row
            
            '计算价税合计
            m_BillInterface.SetFieldValue "FAmountIncludeTaxFor", BoxQty * BoxSize * dblPrice, Row
            m_BillInterface.SetFieldValue "FAmountIncludeTax", BoxQty * BoxSize * dblPrice, Row
            
        
        
            For i = 1 To m_BillInterface.BillEntrys(2).GridMaxDataRowNum
                ItemID = Val(m_BillInterface.GetFieldValue("FProductID2", i))
                If ItemID <> 0 Then
                    Amt = Val(m_BillInterface.GetFieldValue("FAmountIncludeTaxFor", i))
                    SumAmt = SumAmt + Amt
                End If
            Next
            
            m_BillInterface.SetFieldValue "FAmountFor", SumAmt, 1
            m_BillInterface.SetFieldValue "FAmount2", SumAmt, 1
            m_BillInterface.SetFieldValue "FTotalAmount", SumAmt, 1
            m_BillInterface.SetFieldValue "FTotalAmountFor", SumAmt, 1
            
        End If
    End If
    
    If m_BillInterface.FID = 200000018 Then '移仓申请单
    
        If dct("FKey") = "FDCStockID" Then '自动带出默认仓位
        StockID = Val(m_BillInterface.GetFieldValue("FDCStockID", Row))
        
            strSql = "select t2.FDefaultSPID,t3.FNumber from t_Stock t1 inner join t_StockPlaceGroup t2 on t1.FSPGroupID=t2.FSPGroupID "
            strSql = strSql & vbCrLf & "inner join t_StockPlace t3 on t2.FDefaultSPID=t3.FSPID Where t1.FItemID = " & StockID
            
            Set rs = m_BillInterface.K3Lib.GetData(strSql)

                    If rs.Fields("FNumber") = "*" Then

                        Set dctField = m_BillInterface.GetFieldInfoByKey("FDCSPID", "FDCSPID", 1)
                       m_BillInterface.LockCell dctField, Row, True
                    Else
                        m_BillInterface.SetFieldValue "FDCSPID", rs.Fields("FNumber"), Row
                    End If

        End If
        
        If dct("FKey") = "FDecimal2" Then
            BoxQty = Val(m_BillInterface.GetFieldValue("FDecimal2", Row))
            BoxSize = Val(m_BillInterface.GetFieldValue("FBaseProperty2", Row))
            ItemID = Val(m_BillInterface.GetFieldValue("FItemID", Row))
            
            If ItemID <> 0 Then
                Volume = Val(m_BillInterface.GetFieldValue("FBaseProperty", Row))
                Weight = Val(m_BillInterface.GetFieldValue("FBaseProperty1", Row))
                Qty = BoxQty * BoxSize
'                Qty = Val(m_BillInterface.GetFieldValue("FQty", Row))
                BoxModel = Val(m_BillInterface.GetFieldValue("FBaseProperty2", Row))
                m_BillInterface.SetFieldValue "FQty", Qty, Row
                If BoxModel <> 0 Then
                    '计算体积，重量
                    m_BillInterface.SetFieldValue "FDecimal", Qty / BoxModel * Volume, Row
                    m_BillInterface.SetFieldValue "FDecimal1", Qty / BoxModel * Weight, Row
                End If
            End If
        End If
        
        If dct("FKey") = "FBatchNo" Then '计算到期日
            KFDate = m_BillInterface.GetFieldValue("FKFDate", Row)
            KFPeriod = m_BillInterface.GetFieldValue("FKFPeriod", Row)
            PeriodDate = DateAdd("d", KFPeriod, KFDate)
            m_BillInterface.SetFieldValue "FPeriodDate", PeriodDate, Row
        End If
        
    End If
    
End Sub
