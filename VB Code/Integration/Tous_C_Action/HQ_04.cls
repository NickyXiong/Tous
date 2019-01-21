VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HQ_04"
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

    strSQL = "select t1.FPOSNumber,t2.FNumber,sum(t3.FConsignAmount)FAmount,sum(t3.FDiscountAmount)FDiscount  "
    strSQL = strSQL & vbCrLf & "from ICStockBill t1 inner join t_Organization t2 on t1.FSupplyID=t2.FItemID"
    strSQL = strSQL & vbCrLf & "inner join ICStockBillEntry t3 on t1.FInterID=t3.FInterID"
    strSQL = strSQL & vbCrLf & "Where FTranType = 21 And FCheckerID > 0 And IsNull(FIsExportedToHQ, 0) = 0"
    strSQL = strSQL & vbCrLf & "group by t1.FPOSNumber,t2.FNumber"

    Set rs = TB_Runtime.ExecuteSQL(strSQL)
'    Set rs = ExecSQL(strSQL, m_Dsn)
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            Sleep (2000)
                        
            With notDataCLs
                lInstr = InStr(1, strOriFileName, ".")
                strNewFileName = Left(strOriFileName, lInstr - 7) & Format(Time, "hhmmss") & Mid(strOriFileName, lInstr, 4)
                strNewFileName = Replace(strNewFileName, "_", "_" & rs.Fields("FNumber").Value & "_", 1, 1)
                datas.FileName = strNewFileName
                
                
                Set obj = CreateObject("Tous_M_Action.HQ_04")
                billData = obj.ExportData(m_Dsn, retInfo, rs.Fields("FPOSNumber").Value)
                Set obj = Nothing
                
                If TB_Runtime.UboundEx(billData) > -1 Then
                    
                    ExportXML datas.CacheRoot & strNewFileName, billData, CDbl(rs.Fields("FAmount").Value), CDbl(rs.Fields("FDiscount").Value) '导出数据
                    
                    TB_Runtime.K3Log datas, "Exportation is successful!POS No:[" & rs.Fields("FPOSNumber").Value & "]", False
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
    TB_Runtime.K3Log datas, "Tous_C_Action.HQ_04.ITask_Run" & Err.Description, True
End Function

Private Sub ITask_Dispose(param As TB_Context.TBParameters)
    'Set notDataCLs = Nothing
End Sub

Public Function GetDsn(k3svr As TB_Context.TYPE_K3SERVER) As String '组合字符串
    m_Dsn = "Persist Security Info=True;Provider=SQLOLEDB.1;User ID=" & k3svr.DBUsername & ";Password=" & k3svr.DBPassword & ";Data Source=" & k3svr.DBServer & ";Initial Catalog=" & k3svr.DBName
    'k3DsnStr = "ConnectString={" & gConString & "};UserName=" & Set_Biller & ";UserID=" & Set_BillerID & ";DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=gl;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=" & Set_BillerUUID
    GetDsn = m_Dsn
    
End Function

Private Sub ExportXML(strFile As String, billData As Variant, Amount As Double, Discount As Double)
    Dim xmlDocum As MSXML2.DOMDocument
    Dim xmlRoot As MSXML2.IXMLDOMElement
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim xmlElement As MSXML2.IXMLDOMElement
    Dim xmlHeaderElement As MSXML2.IXMLDOMElement
    Dim xmlItemElement As MSXML2.IXMLDOMElement
    
    Dim xmlTaxElement As MSXML2.IXMLDOMElement
    
    Dim pi As IXMLDOMProcessingInstruction
    Dim xmlAttributs As MSXML2.IXMLDOMAttribute
    Dim i As Integer, j As Integer
    
    Dim s() As String, sp() As String
    
    Set xmlDocum = New MSXML2.DOMDocument
    Set xmlRoot = xmlDocum.createElement("NS1:tickets")
    Set xmlDocum.documentElement = xmlRoot
    
    Call xmlRoot.setAttribute("xmlns:NS1", "http://www.tous.com/Tickets")
    
    Set xmlElement = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "ticket", "")
    xmlRoot.appendChild xmlElement
    
    '设置Header
    Set xmlHeaderElement = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "header", "")
    xmlElement.appendChild xmlHeaderElement
    
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "store", "")
    xmlNode.Text = billData(0)(2)
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "date", "")
    xmlNode.Text = billData(0)(1)
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "time", "")
    xmlNode.Text = billData(0)(5)
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "serie", "")
    xmlNode.Text = "1"
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "ticket", "")
    xmlNode.Text = Val(billData(0)(0))
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "currency", "")
    xmlNode.Text = "RMB"
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "POS", "")
    xmlNode.Text = "1"
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "amount", "")
    xmlNode.Text = Format(Round(Amount, 2), "0.00")
    xmlHeaderElement.appendChild xmlNode
    Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "discount", "")
'    xmlNode.Text = IIf(Amount + Discount = 0, 0, Round(100 * Discount / (Amount + Discount), 2))
'    If Amount + Discount = 0 Then
        xmlNode.Text = 0
'    Else
'        xmlNode.Text = Round(100 * Discount / (Amount + Discount), 2)
'    End If
    xmlHeaderElement.appendChild xmlNode
    
    'Add Payment Information
    
'    s = Split(billData(0)(10), ",")
'    For i = 0 To UBound(s)
'        Set xmlItemElement = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "paymentMethods", "")
'        xmlHeaderElement.appendChild xmlItemElement
'
'        sp = Split(Mid(s(i), 2, Len(s(i)) - 2), "/")
'        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "idPaymentMethod", "")
'        xmlNode.Text = "1"
'        xmlItemElement.appendChild xmlNode
'        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "amount", "")
'        xmlNode.Text = sp(2)
'        xmlItemElement.appendChild xmlNode
'    Next
    
    '设置item
    For i = 0 To UBound(billData)
        Set xmlItemElement = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "lines", "")
        xmlElement.appendChild xmlItemElement
        
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "qualificator", "")
        xmlNode.Text = "003"
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "lineNumber", "")
        xmlNode.Text = i + 1
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "sign", "")
        xmlNode.Text = IIf(Val(billData(i)(6)) = 1, "+", "-")
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "type", "")
'        xmlNode.Text = IIf(Val(billData(i)(6)) = 1, 1, 2)
        xmlNode.Text = "2"
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "quantity", "")
        xmlNode.Text = Format(billData(i)(4), "0.0")
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "amount", "")
        xmlNode.Text = Val(billData(i)(9))
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "discount", "")
        xmlNode.Text = Format(Round(billData(i)(8), 2), "0.00")
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "netDiscountAmount", "")
        xmlNode.Text = Format(Round(billData(i)(7), 2), "0.00")
        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "netAmount", "")
        xmlNode.Text = Format(Round(Val(billData(i)(9)) / 1.16, 2), "0.00")
        xmlItemElement.appendChild xmlNode
        
        '处理tax
        Set xmlTaxElement = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "taxes", "")
        xmlItemElement.appendChild xmlTaxElement
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "taxPercentage", "")
        xmlNode.Text = "16.00"
        xmlTaxElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "taxValue", "")
        xmlNode.Text = Format(Round(Val(billData(i)(9)) / 1.16 * 0.16, 2), "0.00")
        xmlTaxElement.appendChild xmlNode
        
        
'        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "vatCode", "")
'        xmlNode.Text = 0
'        xmlItemElement.appendChild xmlNode
        Set xmlNode = xmlDocum.createNode(MSXML2.NODE_ELEMENT, "SKU", "")
        xmlNode.Text = billData(i)(3)
        xmlItemElement.appendChild xmlNode
    Next
    
    Set pi = xmlDocum.createProcessingInstruction("xml", "version='1.0' encoding='utf-8'")
    Call xmlDocum.InsertBefore(pi, xmlDocum.childNodes(0))
    xmlDocum.Save strFile

End Sub





