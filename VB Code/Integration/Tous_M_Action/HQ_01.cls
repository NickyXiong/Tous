VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "HQ_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Dsns As String
Private m_K3Dsn As String
Private m_MDConnection As ADODB.Connection
Private m_strForDict(1 To 14) As Long
Private relation_forDict(0 To 59, 1 To 2) As String

 '字段说明
Private Const CONST_FGarment_chs = 50 'Garment_chs
Private Const CONST_FMaterial_chs = 51 'Material_chs
Private Const CONST_FLining_chs = 52 'Lining_chs
Private Const CONST_FSoleMaterial_chs = 53 'SoleMaterial_chs
Private Const CONST_FCountryOf_chs = 54 'Countryoforigin_chs
Private Const CONST_FSize_chs = 55 'Size_chs
Private Const CONST_FColor_chs = 57 'Color_chs
Private Const CONST_FStandard_chs = 58 'Standard_chs
Private Const CONST_FCHHSCode_chs = 59 'ChinaHSCode_chs

Private fUnnitID As Long '单位
Private FUnitGroupID As Long
Private fAcctID As Long
Private fSaleAcctID As Long
Private fCostAcctID As Long

Private lFMacroFamily As Long
Private lFFamily As Long
Private lFType As Long
Private lFMetal As Long

Private Sub Init()
    Set m_MDConnection = New ADODB.Connection
    With m_MDConnection
        .ConnectionString = m_Dsns
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .CursorLocation = adUseClient
        .Open
    End With
    
    m_K3Dsn = "ConnectString={" & m_Dsns & "};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=cc;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;MachineName=127.0.0.1;K3Version=KUE;UUID=00000000-0000-0000-0000-000000000000"
    
    
    '单位
    fUnnitID = GetFUnitID("PCS", FUnitGroupID)
    fAcctID = GetFAccountID("1405.01")
    fSaleAcctID = GetFAccountID("6001.01")
    fCostAcctID = GetFAccountID("6401.02")
    
    
    lFMacroFamily = 0
    lFFamily = 0
    lFType = 0
    lFMetal = 0
End Sub

Public Function ImportItem(ByVal m_Dsn As String, ByVal dic As KFO.Dictionary, ByRef retInfo As String, ByRef fNumber As String) As Boolean   '保存物料
Dim strSql As String
Dim rs As ADODB.Recordset
On Error GoTo HError
    m_Dsns = m_Dsn
    Init  '初始化
    
    fNumber = CStr(dic("SKU"))
    If fNumber = "" Then
        retInfo = "SKU can not be null"
        Exit Function
    End If
    
    'Add MacroFamily/Family/Type/Metel auto creation - 20190118
    If Len(dic("macrofamily")) > 0 Then
        lFMacroFamily = CreateTousItem(m_Dsn, dic("macrofamily"), 3001)
    End If
    
    If Len(dic("family")) > 0 Then
        lFFamily = CreateTousItem(m_Dsn, dic("family"), 3002)
    End If
    
    If Len(dic("type")) > 0 Then
        lFType = CreateTousItem(m_Dsn, dic("type"), 3003)
    End If
    
    If Len(dic("metal")) > 0 Then
        lFMetal = CreateTousItem(m_Dsn, dic("metal"), 3004)
    End If
    'Add MacroFamily/Family/Type/Metel auto creation - 20190118
    
    
    strSql = "SELECT FItemID FROM t_Item t1 WHERE t1.FItemClassID=4 AND t1.FNumber='" & fNumber & "'"
    Set rs = m_MDConnection.Execute(strSql)
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            ImportItem = UpdateFitem(rs("FItemID"), dic, retInfo)
        Else
            ImportItem = InsertFitem(dic, retInfo)
        End If
    Else
        ImportItem = InsertFitem(dic, retInfo)
    End If
    Set rs = Nothing
    GetObjectContext.SetComplete
    Exit Function
HError:
    GetObjectContext.SetAbort
    ImportItem = False
    retInfo = Err.Description
    '写日志
    
End Function

Public Function UpdateFitem(ByVal FItemID As String, ByVal dic As KFO.Dictionary, ByRef retInfo As String) As Boolean  '更新物料
    Dim strSql As String
    
    strSql = " UPDATE t_Item SET FName='" & CStr(dic("Desc")) & "',FFullName='" & CStr(dic("Desc")) & "'"
    strSql = strSql & " WHERE FItemID =" & FItemID
    m_MDConnection.Execute strSql
    
    strSql = "UPDATE t_ICItem SET "
    strSql = strSql & " FEAN8='" & dic("EAN8") & "',FEAN13='" & dic("EAN13") & "',FHQCurrency='" & dic("Currency") & "', "
    strSql = strSql & " FHQCostPrice=" & Val(dic("HQCostPrice")) & ",FHQRetailPrice='" & Val(dic("HQRetailPrice")) & "', "
    strSql = strSql & " FMacroFamily='" & dic("macrofamily") & "',FFamily='" & dic("family") & "', "
    strSql = strSql & " FTSType='" & dic("type") & "',FMetal='" & dic("metal") & "', "
    
    strSql = strSql & " FMacroFamilyID='" & lFMacroFamily & "',FFamilyID='" & lFFamily & "', "
    strSql = strSql & " FTSTypeID='" & lFType & "',FMetalID='" & lFMetal & "', "
    
    strSql = strSql & " FUpdatedFlag =1 WHERE FItemID =" & FItemID
    m_MDConnection.Execute strSql
    
    UpdateFitem = True
End Function

Public Function InsertFitem(ByRef dic As KFO.Dictionary, ByRef retInfo As String) As Boolean  '新增物料
Dim strSql As String
Dim classItemId As Long
Dim fNumber As String
Dim rs As ADODB.Recordset


Dim strCheckSQL As String
Dim lFBrandID As Long
Dim rsTemp As ADODB.Recordset


    Set par1 = New ADODB.Parameter
    Set par2 = New ADODB.Parameter
    par1.Name = "TableName"
    par1.Value = "t_Item"
    
    par2.Name = "FInterID"
    par2.Value = 0
    
    Set pars = ExceProc("GetICMaxNum", par1, par2, m_MDConnection)
    classItemId = pars("@FInterID").Value '内码
    
    strSql = " INSERT INTO t_Item ([FItemID],[FItemClassID],[FParentID],[FNumber],[FDetail],[FName],[FFullName],[FFullNumber])"
    strSql = strSql & vbCrLf & "VALUES (" & classItemId & ",4,0,'" & CStr(dic("SKU")) & "', 1,'" & CStr(dic("Desc")) & "','"
    strSql = strSql & CStr(dic("Desc")) & "','" & CStr(dic("SKU")) & "')"
    m_MDConnection.Execute strSql
    
    strSql = " INSERT INTO t_ICItem([FItemID],[FModel],[FName],[FShortNumber],[FNumber],[FBrNo]" '6列
    strSql = strSql & vbCrLf & ",[FOrderPrice],[FPerWastage],[FErpClsID],[FUnitID],[FUnitGroupID],[FDefaultLoc]"
    strSql = strSql & vbCrLf & ",[FSource],[FQtyDecimal],[FHighLimit]"
    strSql = strSql & vbCrLf & ",[FFullName],[FOrderUnitID],[FSaleUnitID],[FStoreUnitID]"
    strSql = strSql & vbCrLf & ",[FProductUnitID],[FTypeID],[FOrderRector],[FWWHghPrc],[FSOLowPrc],[FSalePrice]"
    strSql = strSql & vbCrLf & ",[FTrack],[FAPAcctID],[FBookPlan],[FBeforeExpire]"
    strSql = strSql & vbCrLf & ",[FTaxRate],[FCheckCycUnit],[FTotalTQQ],[FBatchAppendQty],[FOrderPoint],[FBatFixEconomy]"
    strSql = strSql & vbCrLf & ",[FProductPrincipal],[FDailyConsume],[FPlanner],[FIsSnManage]"
    
    
    strSql = strSql & vbCrLf & ",[FPOVAcctID],[FPIVAcctID],[FMCVAcctID],[FPCVAcctID],[FSLAcctID],[FCAVAcctID]"
    strSql = strSql & vbCrLf & ",[FProChkMde],[FWWChkMde],[FSOChkMde],[FWthDrwChkMde],[FStkChkMde],[FOtherChkMde]"
    strSql = strSql & vbCrLf & ",[FIdentifier],[FSampStdCritical],[FSampStdStrict],[FSampStdSlight],[FNameEn],[FModelEn]"
    strSql = strSql & vbCrLf & ",[FHSNumber],[FFirstUnitRate],[FSecondUnitRate],[FManageType],[FPlanPrice],[FCtrlType],[FCtrlStraregy]"
    strSql = strSql & vbCrLf & ",[FAcctID],[FSaleAcctID],[FCostAcctID],[FDSManagerID],[FForbbitBarcodeEdit],[FOrderDept],[FAuxInMrpCal]"
    
    
    
    '新增自定义字段
    strSql = strSql & vbCrLf & ",[FMacroFamily],[FFamily],[FTSType],[FMetal]"
    strSql = strSql & vbCrLf & ",[FMacroFamilyID],[FFamilyID],[FTSTypeID],[FMetalID]"
    strSql = strSql & vbCrLf & ",[FUpdatedFlag],[FEAN8],[FEAN13],[FHQCurrency],[FHQCostPrice],[FHQRetailPrice])"


    
    '系统标准字段属性值
    strSql = strSql & vbCrLf & " VALUES(" & classItemId & ",'','" & CStr(dic("Desc")) & "','" & CStr(dic("SKU")) & "','" & CStr(dic("SKU")) & "',0,"
    strSql = strSql & vbCrLf & "0,0,1," & fUnnitID & "," & FUnitGroupID & ",0," '需确定单位资料后补填
    strSql = strSql & vbCrLf & "0,4,1000,"
    strSql = strSql & vbCrLf & "'" & CStr(dic("Desc")) & "'," & fUnnitID & "," & fUnnitID & "," & fUnnitID & "," 'FFullName
    strSql = strSql & vbCrLf & "" & fUnnitID & ",0,0,0,0,0,"
    strSql = strSql & vbCrLf & "76,0,0,0," '需确定单位资料后补填
    strSql = strSql & vbCrLf & "16,0,0,1,0,0," '[FBatchAppendQty]
    strSql = strSql & vbCrLf & "0,0,0,0,"
    
    strSql = strSql & vbCrLf & " 0,0,0,0,0,0,"
    strSql = strSql & vbCrLf & " 352,352,352,352,352,352,"
    strSql = strSql & vbCrLf & " 0,0,0,0,'" & CStr(dic("Desc")) & "','',"
    strSql = strSql & vbCrLf & " 0,0,0,0,0,14039,14041,"
    
    strSql = strSql & fAcctID & "," & fSaleAcctID & "," & fCostAcctID & ",0,0,0,0,"
    
    '新增自定义字段属性值
    strSql = strSql & vbCrLf & "'" & CStr(dic("macrofamily")) & "','" & CStr(dic("family")) & "','" & CStr(dic("type")) & "','" & CStr(dic("metal")) & "',"
    strSql = strSql & vbCrLf & "'" & lFMacroFamily & "','" & lFFamily & "','" & lFType & "','" & lFMetal & "',"
    strSql = strSql & vbCrLf & "1,'" & CStr(dic("EAN8")) & "','" & CStr(dic("EAN13")) & "','" & CStr(dic("Currency")) & "'," & Val(dic("HQCostPrice")) & "," & Val(dic("HQRetailPrice")) & ")"

    
    m_MDConnection.Execute (strSql)
    
    InsertFitem = True
End Function

Private Function GetSubmessageFInterID(ByVal submessageFID As String, ByVal submessageFParentID As String, ByVal submessageFName As String) As String
Dim strSql As String
Dim rs As ADODB.Recordset
    strSql = "SELECT FInterID FROM t_SubMessage WHERE FID='" & submessageFID & "' AND FParentID='" & submessageFParentID & "'"
    If submessageFName <> "" Then
        strSql = strSql & " FName='" & submessageFName & "'"
    End If
    Set rs = m_MDConnection.Execute(strSql)
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then '导出数据
            GetSubmessageFInterID = rs("FInterID")
        Else
            GetSubmessageFInterID = "0"
        End If
    Else
        GetSubmessageFInterID = "0"
    End If
    Set rs = Nothing
End Function

Private Function GetFUnitID(ByVal fUnitName As String, ByRef FUnitGroupID As Long) As Long
Dim strSql As String
Dim rs As ADODB.Recordset
    strSql = "SELECT FMeasureUnitID,FUnitGroupID FROM t_MeasureUnit WHERE FName='" & fUnitName & "'"
    Set rs = m_MDConnection.Execute(strSql)
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then '导出数据
            GetFUnitID = rs("FMeasureUnitID")
            FUnitGroupID = rs("FUnitGroupID")
        Else
            GetFUnitID = 0
            FUnitGroupID = 0
        End If
    Else
        GetFUnitID = 0
        FUnitGroupID = 0
    End If
    Set rs = Nothing
End Function
Private Function GetFAccountID(ByVal fAccountNumber As String) As Long
Dim strSql As String
Dim rs As ADODB.Recordset
    strSql = "SELECT  FAccountID FROM t_Account WHERE FNumber='" & fAccountNumber & "'"
    Set rs = m_MDConnection.Execute(strSql)
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then '导出数据
            GetFAccountID = rs("FAccountID")
        Else
            GetFAccountID = 0
        End If
    Else
        GetFAccountID = 0
    End If
    Set rs = Nothing
End Function

Public Function ExceProc(ByVal procName As String, ByVal par1 As ADODB.Parameter, ByVal par2 As ADODB.Parameter, ByVal con As ADODB.Connection) As ADODB.Parameters
    Dim adoComm As Object
    '// 创建一个对象，我们用来调用存储过程
    Set adoComm = CreateObject("ADODB.Command")
    With adoComm
        '// 设置连接，假设 adoConn 为已经连接的 ADODB.Connection 对象
        .ActiveConnection = con
        '// 类型为存储过程，adCmdStoredProc = 4
        .CommandType = 4
        '// 存储过程名称
        .CommandText = procName
        '// 设置输入参数
        .Parameters(1) = par1
        .Parameters(2) = par2
        '// 执行存储过程
        .Execute
        Set ExceProc = .Parameters
    End With
    '// 释放对象
    Set adoComm = Nothing
End Function

Public Function CreateTousItem(ByVal sDsn As String, ByVal strName As String, ByVal lClassID As Long) As Long
    Dim item As New KFO.Dictionary
    Dim number As String
    Dim strSql As String
    Dim rs As ADODB.Recordset
    Dim rsCount As ADODB.Recordset
On Error Resume Next

    strSql = "SELECT FItemID FROM t_Item_" & CStr(lClassID) & " WHERE FName='" & Trim(strName) & "'"
    Set rs = m_MDConnection.Execute(strSql)
    
    If rs.RecordCount <= 0 Then
        strSql = "select max(FNumber)FNumber from t_Item_" & CStr(lClassID)
        Set rsCount = m_MDConnection.Execute(strSql)
        
        number = Val(rsCount.Fields("FNumber").Value) + 1
        item("FNumber") = number
        item("FFullNumber") = number
        item("FShortNumber") = number
        item("FName") = strName
        
        item("FDetail") = True '如果是明细，则必须是true，如果是上级组，则是false
        item("FLevel") = 1
        item("FParentID") = 0
        
        Dim oSave As Object
        Set oSave = CreateObject("EBSGLItem.ItemUpdate")
        CreateTousItem = oSave.Create(m_K3Dsn, lClassID, item)
    Else
        CreateTousItem = rs.Fields("FItemID").Value
    End If

End Function

