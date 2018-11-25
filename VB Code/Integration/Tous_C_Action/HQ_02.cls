VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HQ_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements ITask
'Private notDataCLs As Object '读写数据类
Private datas As TB_Context.TYPE_ACTION_DATA
Private m_Dsn As String
Private m_K3Connection As ADODB.Connection


Private Sub Class_Initialize()
'Dim k3svr As TB_Context.TYPE_K3SERVER
'Dim param As TB_Context.TBParameters
    'Set k3svr = New TB_Context.TYPE_K3SERVER
    'Set data = New TB_Context.TYPE_ACTION_DATA
    'Set param = New TB_Context.TBParameters
'    datas.FileName = "C:\Tory Burch\exapleDATA\Catalog_201212201731260.csv"
'    datas.CacheRoot = ""
'    k3svr.DBName = "AIS20130115152640"
'    k3svr.DBPassword = "123"
'    k3svr.DBServer = "UNIONWAR-PTKAWT\MSSQL_LJ"
'    k3svr.DBUsername = "sa"
'    k3svr.K3Username = "administrator"
'    ITask_Init k3svr, datas, param
'    ITask_Run param
End Sub

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

    Dim xmlDocument As New DOMDocument
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim temp As IXMLDOMNodeList
    Dim obj As Object
    
    Dim retInfo As String, strFNumber As String
    Dim result As Boolean
    Dim i As Long, j As Long
    Dim strSQL As String
    
    Dim dic As KFO.Dictionary
    Dim vec As KFO.Vector
    
    Dim fso As New FileSystemObject
    
On Error GoTo HERROR

    If ValidateXML(datas.CacheRoot & datas.FileName, "http://www.tous.com/articles/pricelist", datas.CacheRoot & "\xsd\PriceListInformation.xsd", retInfo) = False Then
        ExportErrorXML Replace(datas.CacheRoot, "pending", "error") & "error." & datas.FileName, retInfo
        If fso.FileExists(Replace(datas.CacheRoot, "pending", "error") & datas.FileName) = True Then
            fso.DeleteFile (Replace(datas.CacheRoot, "pending", "error") & datas.FileName)
        End If
        fso.MoveFile datas.CacheRoot & datas.FileName, Replace(datas.CacheRoot, "pending", "error")
        
        TB_Runtime.K3Log datas, "File format error:" & retInfo, True
        Exit Function
    End If
    retInfo = ""
    
    xmlDocument.async = False
    xmlDocument.Load (datas.CacheRoot & datas.FileName)
    
    Set vec = New KFO.Vector
        
    Set xmlNodeList = xmlDocument.getElementsByTagName("priceList") '遍历指定的结点gendata
    For i = 0 To xmlNodeList.length - 1
        Set xmlNode = xmlNodeList.Item(i)
        Set dic = New KFO.Dictionary
        
        'type
        Set temp = xmlNode.selectNodes("type")
        If temp.length <> 0 Then
            dic("Type") = temp.Item(0).Text
        Else
            dic("Type") = ""
        End If
        Set temp = Nothing
        
        'SKU
        Set temp = xmlNode.selectNodes("sku")
        If temp.length <> 0 Then
            dic("SKU") = temp.Item(0).Text
        Else
            dic("SKU") = ""
        End If
        Set temp = Nothing
        
        'value
        Set temp = xmlNode.selectNodes("value")
        If temp.length <> 0 Then
            dic("Value") = temp.Item(0).Text
        Else
            dic("Value") = 0
        End If
        Set temp = Nothing
        
        strFNumber = ""
        Set obj = CreateObject("Tous_M_Action.HQ_02")
        result = obj.UpdatePrice(m_Dsn, dic, retInfo)
        If result = False Then
            TB_Runtime.K3Log datas, "Update Price failed[" & dic("SKU") & "]: " & retInfo, True
        Else
            TB_Runtime.K3Log datas, "Update Price successfully!Article[" & dic("SKU") & "]", False
        End If
    
        Set dic = Nothing
    Next
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlDocument = Nothing
    
    
    Set vec = Nothing
    ITask_Run = True
    Exit Function
HERROR:
    ITask_Run = False
    TB_Runtime.K3Log datas, "Tous_C_Action.HQ_02.ITask_Run" & Err.Description, True
End Function

Private Sub ITask_Dispose(param As TB_Context.TBParameters)
    'Set notDataCLs = Nothing
End Sub

Public Function GetDsn(k3svr As TB_Context.TYPE_K3SERVER) As String '组合字符串
    m_Dsn = "Persist Security Info=True;Provider=SQLOLEDB.1;User ID=" & k3svr.DBUsername & ";Password=" & k3svr.DBPassword & ";Data Source=" & k3svr.DBServer & ";Initial Catalog=" & k3svr.DBName
    'k3DsnStr = "ConnectString={" & gConString & "};UserName=" & Set_Biller & ";UserID=" & Set_BillerID & ";DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=gl;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=" & Set_BillerUUID
    GetDsn = m_Dsn
    
End Function






