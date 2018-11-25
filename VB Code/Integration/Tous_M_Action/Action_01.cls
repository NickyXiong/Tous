VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "Action_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Dsns As String
Private m_K3Dsn As String
Private m_MDConnection As ADODB.Connection
Private m_strForDict(1 To 14) As Long
Private relation_forDict(0 To 59, 1 To 2) As String

 '�ֶ�˵��
Private Const CONST_FGarment_chs = 50 'Garment_chs
Private Const CONST_FMaterial_chs = 51 'Material_chs
Private Const CONST_FLining_chs = 52 'Lining_chs
Private Const CONST_FSoleMaterial_chs = 53 'SoleMaterial_chs
Private Const CONST_FCountryOf_chs = 54 'Countryoforigin_chs
Private Const CONST_FSize_chs = 55 'Size_chs
Private Const CONST_FColor_chs = 57 'Color_chs
Private Const CONST_FStandard_chs = 58 'Standard_chs
Private Const CONST_FCHHSCode_chs = 59 'ChinaHSCode_chs

Private fUnnitID As Long '��λ
Private FUnitGroupID As Long
Private fAcctID As Long
Private fSaleAcctID As Long
Private fCostAcctID As Long

Private Sub Init()
    Set m_MDConnection = New ADODB.Connection
    With m_MDConnection
        .ConnectionString = m_Dsns
        .CommandTimeout = 15
        .ConnectionTimeout = 15
        .CursorLocation = adUseClient
        .Open
    End With
    
    m_K3Dsn = "ConnectString={" & m_Dsns & "};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=cc;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;MachineName=127.0.0.1;K3Version=KUE;UUID=00000000-0000-0000-0000-000000000000"
    
    
    '��λ
    fUnnitID = GetFUnitID("PCS", FUnitGroupID)
    fAcctID = GetFAccountID("1")
    fSaleAcctID = GetFAccountID("1")
    fCostAcctID = GetFAccountID("1")
End Sub

Public Function ImportItem(ByVal m_Dsn As String, ByVal dic As kfo.Dictionary, ByRef retInfo As String, ByRef fNumber As String) As Boolean   '��������
Dim strSQL As String
Dim rs As ADODB.Recordset
On Error GoTo HError
    m_Dsns = m_Dsn
    Init  '��ʼ��
    
    fNumber = CStr(dic("SKU"))
    If fNumber = "" Then
        retInfo = "SKU can not be null"
        Exit Function
    End If
    strSQL = "SELECT FItemID FROM t_Item t1 WHERE t1.FItemClassID=4 AND t1.FNumber='" & fNumber & "'"
    Set rs = m_MDConnection.Execute(strSQL)
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
    SaveMarter = False
    retInfo = "TBM_Action_Tory.Action1:" & Err.Description
    'д��־
    
End Function

Public Function UpdateFitem(ByVal FItemID As String, ByVal dic As kfo.Dictionary, ByRef retInfo As String) As Boolean  '��������
    Dim strSQL As String
    
    strSQL = " UPDATE t_Item SET FName='" & CStr(dic("Desc")) & "',FFullName='" & CStr(dic("Desc")) & "'"
    strSQL = strSQL & " WHERE FItemID =" & FItemID
    m_MDConnection.Execute strSQL
    
    strSQL = "UPDATE t_ICItem SET "
    strSQL = strSQL & " FEAN8='" & dic("EAN8") & "',FEAN13='" & dic("EAN13") & "',FHQCurrency='" & dic("Currency") & "', "
    strSQL = strSQL & " FHQCostPrice=" & Val(dic("HQCostPrice")) & ",FHQRetailPrice='" & Val(dic("HQRetailPrice")) & "', "
    strSQL = strSQL & " FUpdatedFlag =1 WHERE FItemID =" & FItemID
    m_MDConnection.Execute strSQL
    
    UpdateFitem = True
End Function

Public Function InsertFitem(ByRef dic As kfo.Dictionary, ByRef retInfo As String) As Boolean  '��������
Dim strSQL As String
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
    classItemId = pars("@FInterID").Value '����
    
    strSQL = " INSERT INTO t_Item ([FItemID],[FItemClassID],[FParentID],[FNumber],[FDetail],[FName],[FFullName],[FFullNumber])"
    strSQL = strSQL & vbCrLf & "VALUES (" & classItemId & ",4,0,'" & CStr(dic("SKU")) & "', 1,'" & CStr(dic("Desc")) & "','"
    strSQL = strSQL & CStr(dic("Desc")) & "','" & CStr(dic("SKU")) & "')"
    m_MDConnection.Execute strSQL
    
    strSQL = " INSERT INTO t_ICItem([FItemID],[FModel],[FName],[FShortNumber],[FNumber],[FBrNo]" '6��
    strSQL = strSQL & vbCrLf & ",[FOrderPrice],[FPerWastage],[FErpClsID],[FUnitID],[FUnitGroupID],[FDefaultLoc]"
    strSQL = strSQL & vbCrLf & ",[FSource],[FQtyDecimal],[FHighLimit]"
    strSQL = strSQL & vbCrLf & ",[FFullName],[FOrderUnitID],[FSaleUnitID],[FStoreUnitID]"
    strSQL = strSQL & vbCrLf & ",[FProductUnitID],[FTypeID],[FOrderRector],[FWWHghPrc],[FSOLowPrc],[FSalePrice]"
    strSQL = strSQL & vbCrLf & ",[FTrack],[FAPAcctID],[FBookPlan],[FBeforeExpire]"
    strSQL = strSQL & vbCrLf & ",[FTaxRate],[FCheckCycUnit],[FTotalTQQ],[FBatchAppendQty],[FOrderPoint],[FBatFixEconomy]"
    strSQL = strSQL & vbCrLf & ",[FProductPrincipal],[FDailyConsume],[FPlanner]"
    
    
    strSQL = strSQL & vbCrLf & ",[FPOVAcctID],[FPIVAcctID],[FMCVAcctID],[FPCVAcctID],[FSLAcctID],[FCAVAcctID]"
    strSQL = strSQL & vbCrLf & ",[FProChkMde],[FWWChkMde],[FSOChkMde],[FWthDrwChkMde],[FStkChkMde],[FOtherChkMde]"
    strSQL = strSQL & vbCrLf & ",[FIdentifier],[FSampStdCritical],[FSampStdStrict],[FSampStdSlight],[FNameEn],[FModelEn]"
    strSQL = strSQL & vbCrLf & ",[FHSNumber],[FFirstUnitRate],[FSecondUnitRate],[FManageType],[FPlanPrice],[FCtrlType],[FCtrlStraregy]"
    strSQL = strSQL & vbCrLf & ",[FAcctID],[FSaleAcctID],[FCostAcctID],[FDSManagerID],[FForbbitBarcodeEdit],[FOrderDept],[FAuxInMrpCal]"
    
    
    
    '�����Զ����ֶ�
    strSQL = strSQL & vbCrLf & ",[FUpdatedFlag],[FEAN8],[FEAN13],[FHQCurrency],[FHQCostPrice],[FHQRetailPrice])"


    
    'ϵͳ��׼�ֶ�����ֵ
    strSQL = strSQL & vbCrLf & " VALUES(" & classItemId & ",'','" & CStr(dic("Desc")) & "','" & CStr(dic("SKU")) & "','" & CStr(dic("SKU")) & "',0,"
    strSQL = strSQL & vbCrLf & "0,0,1," & fUnnitID & "," & FUnitGroupID & ",0," '��ȷ����λ���Ϻ���
    strSQL = strSQL & vbCrLf & "0,4,1000,"
    strSQL = strSQL & vbCrLf & "'" & CStr(dic("Desc")) & "'," & fUnnitID & "," & fUnnitID & "," & fUnnitID & "," 'FFullName
    strSQL = strSQL & vbCrLf & "" & fUnnitID & ",0,0,0,0,0,"
    strSQL = strSQL & vbCrLf & "78,0,0,0," '��ȷ����λ���Ϻ���
    strSQL = strSQL & vbCrLf & "17,0,0,1,0,0," '[FBatchAppendQty]
    strSQL = strSQL & vbCrLf & "0,0,0,"
    
    strSQL = strSQL & vbCrLf & " 0,0,0,0,0,0,"
    strSQL = strSQL & vbCrLf & " 352,352,352,352,352,352,"
    strSQL = strSQL & vbCrLf & " 0,0,0,0,'" & CStr(dic("Desc")) & "','',"
    strSQL = strSQL & vbCrLf & " 0,0,0,0,0,14039,14041,"
    
    strSQL = strSQL & fAcctID & "," & fSaleAcctID & "," & fCostAcctID & ",0,0,0,0,"
    
    '�����Զ����ֶ�����ֵ
    strSQL = strSQL & vbCrLf & "1,'" & CStr(dic("EAN8")) & "','" & CStr(dic("EAN13")) & "','" & CStr(dic("Currency")) & "'," & Val(dic("HQCostPrice")) & "," & Val(dic("HQRetailPrice")) & ")"

    
    m_MDConnection.Execute (strSQL)
    
    InsertFitem = True
End Function

Private Function GetSubmessageFInterID(ByVal submessageFID As String, ByVal submessageFParentID As String, ByVal submessageFName As String) As String
Dim strSQL As String
Dim rs As ADODB.Recordset
    strSQL = "SELECT FInterID FROM t_SubMessage WHERE FID='" & submessageFID & "' AND FParentID='" & submessageFParentID & "'"
    If submessageFName <> "" Then
        strSQL = strSQL & " FName='" & submessageFName & "'"
    End If
    Set rs = m_MDConnection.Execute(strSQL)
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then '��������
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
Dim strSQL As String
Dim rs As ADODB.Recordset
    strSQL = "SELECT FMeasureUnitID,FUnitGroupID FROM t_MeasureUnit WHERE FName='" & fUnitName & "'"
    Set rs = m_MDConnection.Execute(strSQL)
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then '��������
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
Dim strSQL As String
Dim rs As ADODB.Recordset
    strSQL = "SELECT  FAccountID FROM t_Account WHERE FNumber='" & fAccountNumber & "'"
    Set rs = m_MDConnection.Execute(strSQL)
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then '��������
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
    '// ����һ�����������������ô洢����
    Set adoComm = CreateObject("ADODB.Command")
    With adoComm
        '// �������ӣ����� adoConn Ϊ�Ѿ����ӵ� ADODB.Connection ����
        .ActiveConnection = con
        '// ����Ϊ�洢���̣�adCmdStoredProc = 4
        .CommandType = 4
        '// �洢��������
        .CommandText = procName
        '// �����������
        .Parameters(1) = par1
        .Parameters(2) = par2
        '// ִ�д洢����
        .Execute
        Set ExceProc = .Parameters
    End With
    '// �ͷŶ���
    Set adoComm = Nothing
End Function

