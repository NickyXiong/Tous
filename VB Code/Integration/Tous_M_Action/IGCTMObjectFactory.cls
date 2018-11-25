VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGCTMObjectFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IGCTMObjectFactory
Private Function IGCTMObjectFactory_CreateObject(ByVal par_WSMLNode As IXMLDOMNode) As Object
    Dim node As IXMLDOMNode
    On Error GoTo IGCTMObjectFactoryTrap
    
    Set node = par_WSMLNode.Attributes.getNamedItem("targetClassName")
    Set IGCTMObjectFactory_CreateObject = Nothing
    If Not (node Is Nothing) Then
        Select Case node.nodeValue
     ' 判断要映射成哪个实体类
            Case "User"
                Set IGCTMObjectFactory_CreateObject = New User
        End Select
    End If
Exit Function
