VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Dsns As String
Private m_K3Dsn As String
Private m_MDConnection As ADODB.Connection
Private datasource As CDataSource
Private m_oParse As CParse
Private strLanguage As String

'/* 索引列表
Private Const CONST_MARK As Long = 0 '表头表体区分标志

Private Sub Init()
    Set m_MDConnection = New ADODB.Connection
    With m_MDConnection
        .ConnectionString = m_Dsns
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .CursorLocation = adUseClient
        .Open
    End With
    
'    m_K3Dsn = "ConnectString={" & m_Dsns & "};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=cc;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;MachineName=127.0.0.1;K3Version=KUE;UUID=00000000-0000-0000-0000-000000000000"
    
End Sub

Private Function InitDataEnv(ByVal sToParse As String) As Boolean
    Set m_oParse = New CParse
    If m_oParse.ParseString(sToParse) Then
        Set datasource = New CDataSource
        Set datasource.ParseObject = m_oParse
    Else
        Err.Raise EBS_E_TypeMismatch, "ParseString"
    End If
    strLanguage = datasource.ParseObject.LanguageID
End Function

Public Function SynBill1000005(ByVal sDsn As String, ByVal rowData As Variant, ByRef strBillNo As String, ByRef retInfo As String) As Boolean  '更新物料
Dim strSql As String, strSQLtemp As String
Dim rs As ADODB.Recordset, rsTemp As ADODB.Recordset
Dim oBill As Object, oSave As Object
Dim InterID As Long, billNO As String, result As String, FPeriod As String
Dim funitid As Long, i As Long, page2RowIndex As Long, page3RowIndex As Long
Dim lFItemID As Long, lUnitID As Long, lSCGroupID As Long
Dim fNewInterID As Long, fExpQty As String
Dim bRet As Boolean
Dim lFExpDate As String, lFActDate As String
Dim blnIsSCGrop As Boolean '是否进行仓位管理

Dim dblAmount As Double
Dim CNN As ADODB.Connection

On Error GoTo HError


    InitDataEnv sDsn
    Set CNN = datasource.Connection

    '检查收款单号是否存在，存在则不允许重复导入
    Set rs = CNN.Execute("select 1 from t_RP_NewReceiveBill where FNumber='" & Trim(CStr(rowData(0)(1))) & "'")
    If rs.RecordCount > 0 Then
        retInfo = "Collection Note Number[" & Trim(CStr(rowData(0)(1))) & "] has existed in ERP system."
        GoTo HError
    End If
    
    Set oBill = CreateObject("AcXm_BosPortal.clsBillPackage")
    With oBill
        .UseDefault True
        .NewBill sDsn, 1000005
        
        .Head("FNumber") = Trim(CStr(rowData(0)(1)))
        .Head("FDate") = Trim(CStr(rowData(0)(2)))
        .Head("FFincDate") = Trim(CStr(rowData(0)(2)))
        
        .Head("FExchangeRateType") = 1
        .Head("FExchangeRate") = 1
        .Head("FCurrencyID") = 1
        
        '收款类型，默认 销售回款
        .Head("FBillType") = 1000
        
        '核算项目类型，默认 客户
        .Head("FItemClassID") = 1
        lFItemID = GetItemID(CNN, Trim(CStr(rowData(0)(0))), 1)
        If lFItemID = 0 Then
            retInfo = "Customer Master Data [" & CStr(rowData(0)(0)) & "] is NOT created in ERP system."
            GoTo HError
        End If
        .Head("FCustomer") = lFItemID
        
        
        .Head("FYear") = Year(CDate(rowData(0)(2)))
        .Head("FPeriod") = Month(CDate(rowData(0)(2)))
        .Head("FPre") = 0
        
        dblAmount = 0
        page2RowIndex = 0
        page3RowIndex = 0
        For i = LBound(rowData) To UBound(rowData)
        
'            page3RowIndex = Trim(CStr(rowData(i)(CONST_P3_FEntryID)))
            .AddRow i + 1, 3
'            .Entry("FEntryID", i + 1, 3) = page3RowIndex
            
            lFItemID = GetItemID(CNN, Trim(CStr(rowData(i)(4))), 4, lUnitID)
            If lFItemID = 0 Then
                retInfo = "Material Master Data [" & CStr(rowData(i)(4)) & "] is NOT created in ERP system."
                GoTo HError
            End If
            .Entry("FItemID", i + 1, 3) = lFItemID '产品代码及其相关字段
            
            
            .Entry("FEntryAccountID", i + 1, 3) = GetFAccountID(CNN, "1122.01")
            
            .Entry("FSettleAmountFor_3", i + 1, 3) = CDbl(rowData(i)(5))
            .Entry("FSettleAmount_3", i + 1, 3) = CDbl(rowData(i)(5))

            dblAmount = dblAmount + CDbl(rowData(i)(5))
        Next
        
        .Head("FAmount") = dblAmount
        .Head("FAmountFor") = dblAmount
        
    End With
    
    Set oSave = CreateObject("AcXm_BosPortal.DataPortal")
'    billNO = Trim(CStr(rowData(0)(CONST_FBILLNO)))
    bRet = oSave.Save(sDsn, oBill, fNewInterID, billNO, result)

    If bRet Then
'       自动审核
'        CheckBosSimple sDsn, 1000005, fNewInterID, -1
'        CheckBosSimple sDsn, 1000005, fNewInterID, 1


    End If
        
    Set oSave = Nothing
    Set oBill = Nothing
    Set rs = Nothing
    
    SynBill1000005 = True
    
    GetObjectContext.SetComplete

    Exit Function
HError:
    SynBill1000005 = False
    retInfo = "Generate Collection Note Failed:" & Err.Description & retInfo
    'GetObjectContext.SetAbort
End Function


Private Function GetFAccountID(ByVal CNN As ADODB.Connection, ByVal fAccountNumber As String) As Long
Dim strSql As String
Dim rs As ADODB.Recordset
    strSql = "SELECT  FAccountID FROM t_Account WHERE FNumber='" & fAccountNumber & "'"
    Set rs = CNN.Execute(strSql)
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

Private Function CheckBosSimple(dsn As String, FClassTypeID As Long, FID As Long, level As Long) As Long
    Dim check As Object
    Set check = CreateObject("K3MBOSMultiCheck.cUpdateData")
    Dim lret As Long
    
    CheckBosSimple = check.AutoCheckBill(dsn, FClassTypeID, FID, 1, 0, True, False, level)
    Set check = Nothing
End Function

