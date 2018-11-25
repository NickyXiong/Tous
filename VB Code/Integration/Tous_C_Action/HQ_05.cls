VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HQ_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements TB_Runtime.ITask

Private m_sDsn As String
Private m_sPath As String
Private m_sFileName As String
Private m_sFullPath As String
Private m_sActionNumer As String
Private m_sTaskNumber As String
Private m_data As TB_Context.TYPE_ACTION_DATA
Private Sub ITask_Dispose(param As TB_Context.TBParameters)

End Sub

Private Function ITask_Init(k3svr As TB_Context.TYPE_K3SERVER, data As TB_Context.TYPE_ACTION_DATA, param As TB_Context.TBParameters) As Boolean
    
   On Error GoTo ITask_Init_Error

    m_sPath = data.CacheRoot
    m_sFileName = data.FileName
    m_sFullPath = m_sPath & m_sFileName
    m_sActionNumer = data.ActionNumber
    m_sTaskNumber = data.TaskNumber
    
    m_data = data
    m_sDsn = "ConnectString={" & K3ConnectionString(k3svr) & "};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=BAD160C5-B4DB-4965-A8A7-0D373C885AE6"
    ITask_Init = True
    
   Exit Function

ITask_Init_Error:
    ITask_Init = False
    
    TB_Runtime.Log data.TaskNumber, "Tous_C_Action.HQ_05.ITask_Init:" & Err.Description
    
End Function


Private Function ITask_Run(param As TB_Context.TBParameters) As Boolean
    Dim blnSyn As Boolean
On Error GoTo ITask_Run_Error

    ITask_Run = ImportDeliveries
   
    Exit Function

ITask_Run_Error:

    TB_Runtime.Log m_data.TaskNumber, "Tous_C_Action.HQ_05.ITask_Run:" & Err.Description
End Function

Private Function ImportDeliveries() As Boolean

    Dim xmlDocument As New DOMDocument
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim temp As IXMLDOMNodeList
    Dim obj As Object
    
    Dim retInfo As String, strFNumber As String
    Dim result As Boolean
    Dim i As Long, j As Long
    Dim strSQL As String
    
    Dim fso As New FileSystemObject
    
    Dim dicHeader As KFO.Dictionary, dic As KFO.Dictionary
    Dim vec As KFO.Vector
    
On Error GoTo HERROR

    If ValidateXML(m_data.CacheRoot & m_data.FileName, "http://www.tous.com/deliveries", m_data.CacheRoot & "\xsd\Deliveries.xsd", retInfo) = False Then
        ExportErrorXML Replace(m_data.CacheRoot, "pending", "error") & "error." & m_data.FileName, retInfo
        If fso.FileExists(Replace(m_data.CacheRoot, "pending", "error") & m_data.FileName) = True Then
            fso.DeleteFile (Replace(m_data.CacheRoot, "pending", "error") & m_data.FileName)
        End If
        fso.MoveFile m_data.CacheRoot & m_data.FileName, Replace(m_data.CacheRoot, "pending", "error")
        
        TB_Runtime.K3Log m_data, "File format error:" & retInfo, True
        Exit Function
    End If
    retInfo = ""
    

    xmlDocument.async = False
    xmlDocument.Load (m_data.CacheRoot & m_data.FileName)
    
    '取表头
    Set dicHeader = New KFO.Dictionary
    Set xmlNodeList = xmlDocument.getElementsByTagName("origin") '供应商
    If xmlNodeList.length > 0 Then
        dicHeader("Supplier") = xmlNodeList.Item(0).nodeTypedValue
    Else
        dicHeader("Supplier") = ""
    End If
    Set xmlNodeList = xmlDocument.getElementsByTagName("number") '总部单号
    If xmlNodeList.length > 0 Then
        dicHeader("DocNumber") = xmlNodeList.Item(0).nodeTypedValue
    Else
        dicHeader("DocNumber") = ""
    End If
    Set xmlNodeList = xmlDocument.getElementsByTagName("date") '日期
    If xmlNodeList.length > 0 Then
        dicHeader("Date") = xmlNodeList.Item(0).nodeTypedValue
    Else
        dicHeader("Date") = Now()
    End If
    Set xmlNodeList = xmlDocument.getElementsByTagName("type") '类型：1采购2退货
    If xmlNodeList.length > 0 Then
        dicHeader("Type") = xmlNodeList.Item(0).nodeTypedValue
    Else
        dicHeader("Type") = ""
    End If
    Set xmlNodeList = xmlDocument.getElementsByTagName("destination") '门店
    If xmlNodeList.length > 0 Then
        dicHeader("Dept") = xmlNodeList.Item(0).nodeTypedValue
    Else
        dicHeader("Dept") = ""
    End If
    
    dicHeader("Currency") = "EUR"    '币别:默认欧元
    
    '取表体
    Set vec = New KFO.Vector
    Set xmlNodeList = xmlDocument.getElementsByTagName("item")
    If xmlNodeList.length = 0 Then
        TB_Runtime.K3Log m_data, "Import Deliveries[" & dicHeader("DocNumber") & "] failed: No item exists in the EDI file.", True
    End If
    
    For i = 0 To xmlNodeList.length - 1
        Set xmlNode = xmlNodeList.Item(i)
        Set dic = New KFO.Dictionary
        
        'SKU
        Set temp = xmlNode.selectNodes("sku")
        If temp.length <> 0 Then
            dic("SKU") = temp.Item(0).Text
        Else
            dic("SKU") = ""
        End If
        Set temp = Nothing
        
        'Qty
        Set temp = xmlNode.selectNodes("quantity")
        If temp.length <> 0 Then
            dic("Qty") = Val(temp.Item(0).Text)
        Else
            dic("Qty") = 0
        End If
        Set temp = Nothing
        
        'Price
        Set temp = xmlNode.selectNodes("unitCost")
        If temp.length <> 0 Then
            dic("Price") = Val(temp.Item(0).Text)
        Else
            dic("Price") = 0
        End If
        Set temp = Nothing
        
        'Amount
        Set temp = xmlNode.selectNodes("unitPrice")
        If temp.length <> 0 Then
            dic("Amount") = Val(temp.Item(0).Text)
        Else
            dic("Amount") = 0
        End If
        Set temp = Nothing
                
        vec.Add dic
        Set dic = Nothing
    Next
        
    strFNumber = ""
    Set obj = CreateObject("Tous_M_Action.HQ_05")
    result = obj.ImportDeliveries(m_sDsn, dicHeader, vec, retInfo)
    If result = False Then
        ExportErrorXML Replace(m_data.CacheRoot, "pending", "error") & "error." & m_data.FileName, retInfo
        If fso.FileExists(Replace(m_data.CacheRoot, "pending", "error") & m_data.FileName) = True Then
            fso.DeleteFile (Replace(m_data.CacheRoot, "pending", "error") & m_data.FileName)
        End If
        fso.MoveFile m_data.CacheRoot & m_data.FileName, Replace(m_data.CacheRoot, "pending", "error")
        
        TB_Runtime.K3Log m_data, "Import Deliveries[" & dicHeader("DocNumber") & "] failed:" & retInfo, True
    Else
        TB_Runtime.K3Log m_data, "Import Deliveries[" & dicHeader("DocNumber") & "]  successfully!", False
    End If
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlDocument = Nothing
    Set dicHeader = Nothing
    Set vec = Nothing
    ImportDeliveries = True
    Exit Function
HERROR:
    ImportDeliveries = False
    TB_Runtime.K3Log m_data, "Tous_C_Action.HQ_05.ITask_Run" & Err.Description, True
End Function






