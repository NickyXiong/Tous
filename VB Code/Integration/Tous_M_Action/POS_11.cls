VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************
'Function: Export SKU Master
'Date: 2013/11/08
'*******************************************************
Option Explicit

Private m_Dsns As String
Private m_K3Dsn As String
Private m_MDConnection As ADODB.Connection
Private m_dataFeil(0 To 28) As String

Private Sub Class_Initialize() '初始化,导出字段位置说明
    m_dataFeil(0) = "FNumber"
    m_dataFeil(1) = "FEAN13"
    m_dataFeil(2) = "FName"
    m_dataFeil(3) = "FBrandCode"
    m_dataFeil(4) = "FBrandName"
    m_dataFeil(5) = "FYearCode"
    m_dataFeil(6) = "FYearName"
    m_dataFeil(7) = "FSeasonCode"
    m_dataFeil(8) = "FSeasonName"
    m_dataFeil(9) = "FCateCode"
    m_dataFeil(10) = "FCateName"
    m_dataFeil(11) = "FSubCateCode"
    m_dataFeil(12) = "FSubCateName"
    m_dataFeil(13) = "FHQRetailPrice"
    
    m_dataFeil(14) = "FMacroFamilyNumber"
    m_dataFeil(15) = "FMacroFamilyName"
    m_dataFeil(16) = "FMacroFamilyCHS"
    m_dataFeil(17) = "FFamilyNumber"
    m_dataFeil(18) = "FFamilyName"
    m_dataFeil(19) = "FFamilyCHS"
    m_dataFeil(20) = "FTSTypeNumber"
    m_dataFeil(21) = "FTSTypeName"
    m_dataFeil(22) = "FTSTypeCHS"
    m_dataFeil(23) = "FMetalNumber"
    m_dataFeil(24) = "FMetalName"
    m_dataFeil(25) = "FMetalCHS"
    m_dataFeil(26) = "FCollectionNumber"
    m_dataFeil(27) = "FCollectionName"
    m_dataFeil(28) = "FCollectionCHS"
End Sub
Private Sub Init() '初始化
    Set m_MDConnection = New ADODB.Connection
    With m_MDConnection
        .ConnectionString = m_Dsns
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .CursorLocation = adUseClient
        .Open
    End With
    
    m_K3Dsn = "ConnectString={" & m_Dsns & "};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=cc;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;MachineName=127.0.0.1;K3Version=KUE;UUID=00000000-0000-0000-0000-000000000000"
End Sub
Public Function ExportData(ByVal m_Dsn As String, ByRef retInfo As String) As Variant '更新物料
Dim rowData() As String
Dim strSql As String
Dim rs As ADODB.Recordset
Dim m_billData As Variant
Dim n As Long, i As Long

On Error GoTo HError
    m_Dsns = m_Dsn
    Init '初始化
'    strSql = "select FNumber,FEAN13,FName,'Tous'FBrandCode,'Tous'FBrandName,''FYearCode,''FYearName,''FSeasonCode,''FSeasonName,"
'    strSql = strSql & vbCrLf & "''FCateCode,''FCateName,''FSubCateCode,''FSubCateName,FHQRetailPrice,FMacroFamily,FFamily,FTSType,FMetal,''FCollection from t_icitem where isnull(FUpdatedFlag,0)=1"

    'Add category - 20190118
    strSql = "select t1.FNumber,FEAN13,t1.FName,'Tous'FBrandCode,'Tous'FBrandName,''FYearCode,''FYearName,''FSeasonCode,''FSeasonName,"
    strSql = strSql & vbCrLf & "''FCateCode,''FCateName,''FSubCateCode,''FSubCateName,FHQRetailPrice,"
    strSql = strSql & vbCrLf & "isnull(t2.FNumber,'') FMacroFamilyNumber,isnull(t2.FName,'') FMacroFamilyName,isnull(t2.F_101,'') FMacroFamilyCHS,"
    strSql = strSql & vbCrLf & "isnull(t3.FNumber,'') FFamilyNumber,isnull(t3.FName,'') FFamilyName,isnull(t3.F_101,'') FFamilyCHS,"
    strSql = strSql & vbCrLf & "isnull(t4.FNumber,'') FTSTypeNumber,isnull(t4.FName,'') FTSTypeName,isnull(t4.F_101,'') FTSTypeCHS,"
    strSql = strSql & vbCrLf & "isnull(t5.FNumber,'') FMetalNumber,isnull(t5.FName,'') FMetalName,isnull(t5.F_101,'') FMetalCHS,"
    strSql = strSql & vbCrLf & "isnull(t6.FNumber,'') FCollectionNumber,isnull(t6.FName,'') FCollectionName,isnull(t6.F_101,'') FCollectionCHS"
    strSql = strSql & vbCrLf & "from t_icitem t1 left join t_Item_3001 t2 on t1.FMacroFamilyID=t2.FItemID"
    strSql = strSql & vbCrLf & "left join t_Item_3002 t3 on t1.FFamilyID=t3.FItemID"
    strSql = strSql & vbCrLf & "left join t_Item_3003 t4 on t1.FTSTypeID=t4.FItemID"
    strSql = strSql & vbCrLf & "left join t_Item_3004 t5 on t1.FMetalID=t5.FItemID"
    strSql = strSql & vbCrLf & "left join t_Item_3005 t6 on t1.FCollectionID=t6.FItemID"
    strSql = strSql & vbCrLf & "Where IsNull(FUpdatedFlag, 0) = 1"

    Set rs = ExecSQL1(strSql, m_K3Dsn)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then '导出数据
            ReDim m_billData(rs.RecordCount - 1)
            ReDim rowData(UBound(m_dataFeil))
            For n = 1 To rs.RecordCount
                'rowdata = Nothing
                For i = 0 To UBound(m_dataFeil) '6列导出数据
                    rowData(i) = CStr(CNulls(rs(m_dataFeil(i)), ""))
                Next i
                m_billData(n - 1) = rowData
                rs.MoveNext
            Next
        End If
    End If
    ExportData = m_billData
    
    strSql = "update t_icitem set FUpdatedFlag=0 where isnull(FUpdatedFlag,0)=1"
    m_MDConnection.Execute strSql
    
    Set rs = Nothing
    Exit Function
HError:
    retInfo = "Tous_M_Action.POS_11.ExportData:" & Err.Description
    '写日志
End Function





