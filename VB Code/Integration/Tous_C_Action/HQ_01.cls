VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HQ_01"
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
    Dim temp As IXMLDOMNodeList, tempDesc As IXMLDOMNodeList
    Dim obj As Object
    
    
    Dim xmlDescList As IXMLDOMNodeList
    Dim xmlParentNode As IXMLDOMNode, xmlDescNode As IXMLDOMNode
    
    Dim retInfo As String, strFNumber As String
    Dim result As Boolean
    Dim i As Long, j As Long
    Dim strSQL As String
    
    Dim dic As KFO.Dictionary
    Dim vec As KFO.Vector
    
    Dim fso As New FileSystemObject
    
On Error GoTo HERROR

'    If ValidateXML(datas.CacheRoot & datas.FileName, "http://www.tous.com/articles", datas.CacheRoot & "\xsd\ArticlesInformation.xsd", retInfo) = False Then
'        ExportErrorXML Replace(datas.CacheRoot, "pending", "error") & "error." & datas.FileName, retInfo
'        If fso.FileExists(Replace(datas.CacheRoot, "pending", "error") & datas.FileName) = True Then
'            fso.DeleteFile (Replace(datas.CacheRoot, "pending", "error") & datas.FileName)
'        End If
'        fso.MoveFile datas.CacheRoot & datas.FileName, Replace(datas.CacheRoot, "pending", "error")
'
'        TB_Runtime.K3Log datas, "File format error:" & retInfo, True
'        Exit Function
'    End If
    retInfo = ""

    xmlDocument.async = False
    xmlDocument.Load (datas.CacheRoot & datas.FileName)
    
    Set vec = New KFO.Vector
        
    Set xmlNodeList = xmlDocument.getElementsByTagName("gendata") '遍历指定的结点gendata
    For i = 0 To xmlNodeList.length - 1
        Set xmlNode = xmlNodeList.Item(i)
        Set dic = New KFO.Dictionary
                        
        'SKU
        Set temp = xmlNode.selectNodes("SKU")
        If temp.length <> 0 Then
            dic("SKU") = temp.Item(0).Text
            strFNumber = temp.Item(0).Text
        Else
            dic("SKU") = ""
        End If
        Set temp = Nothing
        
        'English Description
        'Updated 2018-11-03: 取多语言描述中的英文描述
'        Set temp = xmlNode.selectNodes("obsr")
'        If temp.length <> 0 Then
'            dic("Desc") = temp.Item(0).Text
'        Else
'            dic("Desc") = strFNumber
'        End If
'        Set temp = Nothing
        
        'EAN8
        Set temp = xmlNode.selectNodes("EAN8")
        If temp.length <> 0 Then
            dic("EAN8") = temp.Item(0).Text
        Else
            dic("EAN8") = ""
        End If
        Set temp = Nothing
        
        'EAN13
        Set temp = xmlNode.selectNodes("EAN13")
        If temp.length <> 0 Then
            dic("EAN13") = temp.Item(0).Text
        Else
            dic("EAN13") = ""
        End If
        Set temp = Nothing
        
        'Currency
        Set temp = xmlNode.selectNodes("divisa_coste")
        If temp.length <> 0 Then
            dic("Currency") = temp.Item(0).Text
        Else
            dic("Currency") = ""
        End If
        Set temp = Nothing
        
        'HQ Cost Price
        Set temp = xmlNode.selectNodes("pcov")
        If temp.length <> 0 Then
            dic("HQCostPrice") = Val(temp.Item(0).Text)
        Else
            dic("HQCostPrice") = 0
        End If
        Set temp = Nothing
        
        'HQ Cost Price
        Set temp = xmlNode.selectNodes("pven")
        If temp.length <> 0 Then
            dic("HQRetailPrice") = Val(temp.Item(0).Text)
        Else
            dic("HQRetailPrice") = 0
        End If
        Set temp = Nothing
        
        'Updated 2018-11-03: Add [marcofamily/family/type/metal]
        'marcofamily
        Set temp = xmlNode.selectNodes("macrofamily")
        If temp.length <> 0 Then
            dic("macrofamily") = Replace(temp.Item(0).Text, "'", "''")
        Else
            dic("macrofamily") = ""
        End If
        Set temp = Nothing
        
        'family
        Set temp = xmlNode.selectNodes("family")
        If temp.length <> 0 Then
            dic("family") = Replace(temp.Item(0).Text, "'", "''")
        Else
            dic("family") = ""
        End If
        Set temp = Nothing
        
        'type
        Set temp = xmlNode.selectNodes("type")
        If temp.length <> 0 Then
            dic("type") = Replace(temp.Item(0).Text, "'", "''")
        Else
            dic("type") = ""
        End If
        Set temp = Nothing
        
        'metal
        Set temp = xmlNode.selectNodes("metal")
        If temp.length <> 0 Then
            dic("metal") = Replace(temp.Item(0).Text, "'", "''")
        Else
            dic("metal") = ""
        End If
        Set temp = Nothing
        'Updated 2018-11-03: Add [marcofamily/family/type/metal]
        
        'English Description
        'Updated 2018-11-03: 取多语言描述中的英文描述
        dic("Desc") = ""
        Set xmlParentNode = xmlNode.parentNode
        Set xmlDescList = xmlParentNode.selectNodes("mds/md")  '遍历指定的结点md
        For j = 0 To xmlDescList.length - 1
            Set xmlDescNode = xmlDescList.Item(j)
            
            If xmlDescNode.selectNodes("langKey").Item(0).Text = "en" Then
                Set temp = xmlDescNode.selectNodes("matDsc")
                If temp.length <> 0 Then
                    dic("Desc") = Replace(temp.Item(0).Text, "'", "''")
                Else
                    dic("Desc") = ""
                End If
                Set temp = Nothing
            End If
        Next
        If dic("Desc") = "" Then
            dic("Desc") = strFNumber
        End If
        'English Description
        'Updated 2018-11-03: 取多语言描述中的英文描述
        
        strFNumber = ""
        Set obj = CreateObject("Tous_M_Action.HQ_01")
        result = obj.ImportItem(m_Dsn, dic, retInfo, strFNumber)
        If result = False Then
            TB_Runtime.K3Log datas, "Import Article[" & strFNumber & "] failed:" & retInfo, True
        Else
            TB_Runtime.K3Log datas, "Import Article[" & strFNumber & "] successfully!", False
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
    TB_Runtime.K3Log datas, "Tous_C_Action.HQ_01.ITask_Run" & Err.Description, True
End Function

Private Sub ITask_Dispose(param As TB_Context.TBParameters)
    'Set notDataCLs = Nothing
End Sub

Public Function GetDsn(k3svr As TB_Context.TYPE_K3SERVER) As String '组合字符串
    m_Dsn = "Persist Security Info=True;Provider=SQLOLEDB.1;User ID=" & k3svr.DBUsername & ";Password=" & k3svr.DBPassword & ";Data Source=" & k3svr.DBServer & ";Initial Catalog=" & k3svr.DBName
    'k3DsnStr = "ConnectString={" & gConString & "};UserName=" & Set_Biller & ";UserID=" & Set_BillerID & ";DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=gl;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=" & Set_BillerUUID
    GetDsn = m_Dsn
    
End Function





