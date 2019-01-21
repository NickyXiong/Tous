VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub CallBosBill(ByVal lClassTypeID As Long)
    Dim oDataSrv As Object
    Dim obj As Object
    
    Set oDataSrv = CreateObject("K3ClassTpl.DataSrv")
    oDataSrv.ClassTypeID = lClassTypeID
    
    Set obj = CreateObject("K3SingleClassBill.ClassInfoManager")
    obj.dataSrv = oDataSrv
    obj.Show
    
End Sub

Public Sub CallBosList(ByVal lClassTypeID As Long _
    , Optional ByVal bShowFilter As Boolean = True _
    , Optional ByVal sHeadSort As String = "" _
    , Optional ByVal sEntrySort As String = "" _
    , Optional ByVal sSelPage As String = "" _
    , Optional ByVal sFilter As String = "")
    Dim oDataSrv As Object
    Dim obj As Object
    Dim dctFilter As KFO.Dictionary
    Dim dctControl As KFO.Dictionary
    
    Set oDataSrv = CreateObject("K3ClassTpl.DataSrv")
    oDataSrv.ClassTypeID = lClassTypeID
    
    Set dctControl = New KFO.Dictionary
    dctControl("ShowScheme") = True             ' 选新单要显示过滤方案选择/保存界面
    If sSelPage = "" Then
        sSelPage = ",1,2,3"
    End If
    dctControl("SelPage") = sSelPage         ' 默认选择的页码，不允许修改：",1,3,"
    dctControl("IsLinkBill") = True             '选单调用过滤条件
    oDataSrv.FilterControl = dctControl
    Set dctControl = Nothing
    
    If Not bShowFilter Then
        Set dctFilter = New KFO.Dictionary
        dctFilter("FilterString") = sFilter
        dctFilter("HeadSort") = sHeadSort
        dctFilter("EntrySort") = sEntrySort
        dctFilter("SelectPage") = sSelPage
        oDataSrv.Filter = dctFilter
        Set dctFilter = Nothing
    Else
        If Not oDataSrv.showfilter Then GoTo HExit
    End If
    
    Set obj = CreateObject("K3ClassLookup.BillLookup")
    Set obj.dataSrv = oDataSrv
    obj.ClassTypeID = lClassTypeID
    obj.Show

HExit:
End Sub

Public Function LookUpBosList(ByVal lClassTypeID As Long _
    , Optional ByVal bShowFilter As Boolean = True _
    , Optional ByVal sHeadSort As String = "" _
    , Optional ByVal sEntrySort As String = "" _
    , Optional ByVal sSelPage As String = "" _
    , Optional ByVal sFilter As String = "" _
    , Optional ByVal bSelWhole As Boolean = False) As KFO.Vector
    Dim oDataSrv As Object
    Dim obj As Object
    Dim dctFilter As KFO.Dictionary
    Dim dctControl As KFO.Dictionary
    
    Set oDataSrv = CreateObject("K3ClassTpl.DataSrv")
    oDataSrv.ClassTypeID = lClassTypeID
    
    Set dctControl = New KFO.Dictionary
    dctControl("ShowScheme") = True             ' 选新单要显示过滤方案选择/保存界面
    If sSelPage = "" Then
        sSelPage = ",1,2,3"
    End If
    dctControl("SelPage") = sSelPage         ' 默认选择的页码，不允许修改：",1,3,"
    dctControl("IsLinkBill") = True             '选单调用过滤条件
    oDataSrv.FilterControl = dctControl
    Set dctControl = Nothing
    
    If Not bShowFilter Then
        Set dctFilter = New KFO.Dictionary
        dctFilter("FilterString") = sFilter
        dctFilter("HeadSort") = sHeadSort
        dctFilter("EntrySort") = sEntrySort
        dctFilter("SelectPage") = sSelPage
        oDataSrv.Filter = dctFilter
        Set dctFilter = Nothing
    Else
        If Not oDataSrv.showfilter Then GoTo HExit
    End If
    
    Set obj = CreateObject("K3ClassLookup.BillLookup")
    Set obj.dataSrv = oDataSrv
    obj.ClassTypeID = lClassTypeID
    obj.Lookup
    Set LookUpBosList = obj.VectReturn
HExit:
End Function

Public Function LookUpBosListNoFilterUI(ByVal lClassTypeID As Long, ByVal sFilter As String) As KFO.Vector
    Dim oDataSrv As Object
    Dim obj As Object
    Dim dctFilter As KFO.Dictionary
    Dim dctControl As KFO.Dictionary
    
    Set oDataSrv = CreateObject("K3ClassTpl.DataSrv")
    oDataSrv.ClassTypeID = lClassTypeID
    
    Set dctFilter = New KFO.Dictionary
'    dctFilter("Filterstring") = sFilter
    dctFilter("HeadSort") = ""
    dctFilter("EntrySort") = ""
    dctFilter("SelectPage") = ",1,2,3"
    oDataSrv.Filter = dctFilter
    Set dctFilter = Nothing
    
    Set dctControl = New KFO.Dictionary
    dctControl("ShowScheme") = False             ' 选新单要显示过滤方案选择/保存界面
    dctControl("SelPage") = ",1,2,3"         ' 默认选择的页码，不允许修改：",1,3,"
    dctControl("IsLinkBill") = False             '选单调用过滤条件
    oDataSrv.FilterControl = dctControl
    Set dctControl = Nothing
    
    Set obj = CreateObject("K3ClassLookup.BillLookup")
    Set obj.dataSrv = oDataSrv
    obj.ClassTypeID = lClassTypeID
    obj.Lookup
    Set LookUpBosListNoFilterUI = obj.VectReturn
End Function

Public Sub CallBosBaseList(ByVal lClassTypeID As Long)
    Dim obj As Object
    
    Set obj = CreateObject("K3ClassLookUp.SingleBillLookUp")
    With obj
        
        .ClassTypeID = lClassTypeID
        .classTypeFilter = " FBillTypeID=1 AND FID>0 "
        .ShowType = 0 'Enu_ShowType.Enu_Explore
        .Show
    End With
    Set obj = Nothing

End Sub

Public Function LoadBill(ByVal ClassTypeID As Long, Optional ByVal billID As Long, Optional ByVal billNo As String)
    Dim obj As Object
    Dim dataSrv As Object
    Dim dctData As Object
        
    Set dataSrv = CreateObject("K3ClassTpl.DataSrv")
    dataSrv.ClassTypeID = ClassTypeID
    
    Set obj = CreateObject("K3SingleClassBill.ClassInfoManager")
    With obj
      
        .dataSrv = dataSrv
        If IsMissing(billID) And IsMissing(billNo) Then
            Err.Raise -1, "", "未指定单据编号或单据内码"
        ElseIf IsMissing(billID) Then
            billID = GetBillID(ClassTypeID, billNo)
        End If
        Set dctData = dataSrv.LoadBill(billID)
        .ISNetControl = False
        .ISCheckRight = False
        .Show billID
    End With
    Set obj = Nothing
    Set dataSrv = Nothing
    Set dctData = Nothing
End Function

Private Function GetBillID(ByVal ClassTypeID As Long, ByVal billNo As String)
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim sBillNoKey As String
    Dim sTableName As String
    Dim sPrimaryKey As String
    
    If ClassTypeID = 0 Then Exit Function
    
    strSQL = "select * from ICClassTableInfo where FClassTypeID=" & ClassTypeID
    Set rs = ExecuteSQL(strSQL)
    If rs Is Nothing Then Exit Function
    If rs.RecordCount = 0 Then Exit Function
    rs.Filter = "FKeyWord='BILLNO'"
    If rs.RecordCount > 0 Then
        sBillNoKey = rs.Fields("FFieldName")
        sTableName = rs.Fields("FTableName")
    End If
    rs.Filter = "FKeyWord='Primary'"
    If rs.RecordCount > 0 Then
        sPrimaryKey = rs.Fields("FFieldName")
    End If
    strSQL = "select * from " & sTableName & " where " & sBillNoKey & " =" & "'" & billNo & "'" & vbCrLf
    Set rs = ExecuteSQL(strSQL)
    If rs Is Nothing Then Exit Function
    If rs.RecordCount = 0 Then Exit Function
    
    GetBillID = rs.Fields(sPrimaryKey)
End Function
