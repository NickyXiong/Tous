VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HQ_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements ITask
'Private notDataCLs As Object '读写数据类
Private datas As TB_Context.TYPE_ACTION_DATA
Private m_Dsn As String
Private m_K3Connection As ADODB.Connection


Private Function ITask_Init(k3svr As TB_Context.TYPE_K3SERVER, data As TB_Context.TYPE_ACTION_DATA, param As TB_Context.TBParameters) As Boolean
On Error GoTo HERROR
    GetDsn k3svr
    datas = data
    'Dim notDataCLs As Noteset
    'notDataCLs.SaveData "sd"
    ITask_Init = True
    Exit Function
HERROR:
    ITask_Init = False
    TB_Runtime.Log data.TaskNumber, Err.Description
End Function

Private Function ITask_Run(param As TB_Context.TBParameters) As Boolean
Dim retInfo As String '返回结果信息
Dim rowData() As String '行数据
Dim billData As Variant '单据数据
Dim notDataCLs As Noteset '读写数据类
Dim obj As Object
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim i As Integer, lInstr As Long
Dim strOriFileName As String
Dim strNewFileName As String
On Error GoTo HERROR

    strOriFileName = datas.FileName
    
'    strSQL = "select t1.FInterID,t1.FBillNo,t4.FNumber,t6.FNumber FCurrency,sum(t2.FQty*t5.FAuxPrice) FAmount"
'    strSQL = strSQL & vbCrLf & "from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
'    strSQL = strSQL & vbCrLf & "inner join t_Department t4 on t1.FDeptID=t4.FItemID"
'    strSQL = strSQL & vbCrLf & "inner join POOrder t3 on t2.FSourceInterId=t3.FInterID"
'    strSQL = strSQL & vbCrLf & "inner join POOrderEntry t5 on t5.FInterID=t3.FInterID and t2.FSourceInterId=t5.FInterID and t2.FSourceEntryID=t5.FEntryID"
'    strSQL = strSQL & vbCrLf & "inner join t_Currency t6 on t3.FCurrencyID=t6.FCurrencyID"
'    strSQL = strSQL & vbCrLf & "Where t1.FTranType in(1,10) and FROB=1 And t1.FCheckerID > 0 And IsNull(FIsExportedToHQ, 0) = 0 and isnull(FHQOrigin,'')=''"
'    strSQL = strSQL & vbCrLf & "group by t1.FInterID,t1.FBillNo,t4.Fnumber,t6.FNumber"
    
    strSQL = "select t1.FInterID,t1.FBillNo,t4.FNumber,'RMB' FCurrency,sum(t2.FQty*t2.FAuxPrice) FAmount"
    strSQL = strSQL & vbCrLf & "from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
    strSQL = strSQL & vbCrLf & "inner join t_Department t4 on t1.FDeptID=t4.FItemID"
    strSQL = strSQL & vbCrLf & "left join POOrder t3 on t2.FSourceInterId=t3.FInterID"
    strSQL = strSQL & vbCrLf & "left join POOrderEntry t5 on t5.FInterID=t3.FInterID and t2.FSourceInterId=t5.FInterID and t2.FSourceEntryID=t5.FEntryID"
    strSQL = strSQL & vbCrLf & "left join t_Currency t6 on t3.FCurrencyID=t6.FCurrencyID"
    strSQL = strSQL & vbCrLf & "Where t1.FTranType in(1) and FROB=1 And t1.FCheckerID > 0 and isnull(FHQOrigin,'')='' And IsNull(FIsExportedToHQ, 0) = 0"
    strSQL = strSQL & vbCrLf & "group by t1.FInterID,t1.FBillNo,t4.Fnumber,t6.FNumber"
    strSQL = strSQL & vbCrLf & "Union"
    strSQL = strSQL & vbCrLf & "select t1.FInterID,t1.FBillNo,t4.FNumber,'RMB' FCurrency,sum(t2.FQty*t2.FAuxPrice) FAmount"
    strSQL = strSQL & vbCrLf & "from ICStockBill t1 inner join ICStockBillEntry t2 on t1.FInterID=t2.FInterID"
    strSQL = strSQL & vbCrLf & "inner join t_Department t4 on t1.FDeptID=t4.FItemID"
    strSQL = strSQL & vbCrLf & "inner join POOrder t3 on t2.FSourceInterId=t3.FInterID"
    strSQL = strSQL & vbCrLf & "inner join POOrderEntry t5 on t5.FInterID=t3.FInterID and t2.FSourceInterId=t5.FInterID and t2.FSourceEntryID=t5.FEntryID"
    strSQL = strSQL & vbCrLf & "inner join t_Currency t6 on t3.FCurrencyID=t6.FCurrencyID"
    strSQL = strSQL & vbCrLf & "Where t1.FTranType in(10) and FROB=1 And t1.FCheckerID > 0 and isnull(FHQOrigin,'')='' And IsNull(FIsExportedToHQ, 0) = 0"
    strSQL = strSQL & vbCrLf & "group by t1.FInterID,t1.FBillNo,t4.Fnumber,t6.FNumber"
    

    Set rs = TB_Runtime.ExecuteSQL(strSQL)
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            Sleep (2000)
                            
            With notDataCLs
                lInstr = InStr(1, strOriFileName, ".")
                strNewFileName = Left(strOriFileName, lInstr - 7) & Format(Time, "hhmmss") & Mid(strOriFileName, lInstr, 4)
                strNewFileName = Replace(strNewFileName, "_", "_" & rs.Fields("FNumber").Value & "_", 1, 1)
                datas.FileName = strNewFileName
                
                Set obj = CreateObject("Tous_M_Action.HQ_10")
                billData = obj.ExportData(m_Dsn, retInfo, rs.Fields("FInterID").Value)
                Set obj = Nothing
                
                If TB_Runtime.UboundEx(billData) > -1 Then
                    
                    ExportXML datas.CacheRoot & strNewFileName, billData, rs.Fields("FCurrency").Value, rs.Fields("FAmount").Value '导出数据
                    
                    TB_Runtime.K3Log datas, "Exportation is successful!Bill No:[" & rs.Fields("FBillNo").Value & "]", False
                End If
            End With
            Set notDataCLs = Nothing
            
            rs.MoveNext
        Next
    End If
    
    ITask_Run = True
    Exit Function
HERROR:
    ITask_Run = False
    TB_Runtime.K3Log datas, "Tous_C_Action.HQ_10.ITask_Run" & Err.Description, True
End Function

Private Sub ITask_Dispose(param As TB_Context.TBParameters)
    'Set notDataCLs = Nothing
End Sub

Public Function GetDsn(k3svr As TB_Context.TYPE_K3SERVER) As String '组合字符串
    m_Dsn = "Persist Security Info=True;Provider=SQLOLEDB.1;User ID=" & k3svr.DBUsername & ";Password=" & k3svr.DBPassword & ";Data Source=" & k3svr.DBServer & ";Initial Catalog=" & k3svr.DBName
    'k3DsnStr = "ConnectString={" & gConString & "};UserName=" & Set_Biller & ";UserID=" & Set_BillerID & ";DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=gl;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=" & Set_BillerUUID
    GetDsn = m_Dsn
    
End Function

Private Sub ExportXML(ByVal strFile As String, ByVal billData As Variant, ByVal FCurrency As String, ByVal FAmount As Double)
    Dim xmlDocum As MSXML2.DOMDocument
    Dim xmlRoot As MSXML2.IXMLDOMElement
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim xmlElement As MSXML2.IXMLDOMElement
    Dim xmlHeaderElement As MSXML2.IXMLDOMElement
    Dim xmlItemElement As MSXML2.IXMLDOMElement
    Dim pi As IXMLDOMProcessingInstruction
    Dim xmlAttributs As MSXML2.IXMLDOMAttribute
    Dim i As Integer
    
    Set xmlDocum = New MSXML2.DOMDocument
    Set xmlRoot = xmlDocum.createElement("tns:purchases")
    Set xmlDocum.documentElement = xmlRoot
    
    Call xmlRoot.setAttribute("xmlns:tns", "http://www.tous.com/localPurchases")
    
    Set xmlElement = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "localPurchase", "")
    xmlRoot.appendChild xmlElement
    
    '设置Header
    Set xmlHeaderElement = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "header", "")
    xmlElement.appendChild xmlHeaderElement
    
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "store", "")
    xmlNode.Text = Replace(billData(0)(2), "-V", "")
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "date", "")
    xmlNode.Text = billData(0)(1)
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "time", "")
    xmlNode.Text = Format(Now, "hh:mm:ss")
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "order", "")
    xmlNode.Text = billData(0)(0)
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "vendor", "")
    xmlNode.Text = billData(0)(11)
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "vendorOrder", "")
    xmlNode.Text = billData(0)(0)
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "amount", "")
    xmlNode.Text = CDbl(FAmount)
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "currency", "")
    xmlNode.Text = FCurrency
    xmlHeaderElement.appendChild xmlNode
    
    '设置item
    For i = 0 To UBound(billData)
        Set xmlItemElement = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "item", "")
        xmlElement.appendChild xmlItemElement
        
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "storeDest", "")
        xmlNode.Text = Replace(billData(0)(2), "-V", "")
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "lineNum", "")
        xmlNode.Text = billData(i)(3)
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "ean8", "")
        xmlNode.Text = billData(i)(5)
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "ean13", "")
        xmlNode.Text = billData(i)(6)
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "sku", "")
        xmlNode.Text = billData(i)(4)
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "quantity", "")
        xmlNode.Text = billData(i)(7)
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "unitPrice", "")
        xmlNode.Text = billData(i)(8)
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "unitPriceType", "")
        xmlNode.Text = 1
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "lineAmount", "")
        xmlNode.Text = billData(i)(9)
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "costPrice", "")
        xmlNode.Text = billData(i)(10)
        xmlItemElement.appendChild xmlNode
        
    Next
    
    Set pi = xmlDocum.createProcessingInstruction("xml", "version='1.0' encoding='utf-8'")
    Call xmlDocum.InsertBefore(pi, xmlDocum.childNodes(0))
    xmlDocum.Save strFile

End Sub






