VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "Action_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private datasource As CDataSource
Private m_oParse As CParse
Private strLanguage As String


'/* 包装信息导入生成外购入库
Private Const CONST_A01_FBillNo As Long = 0 '产品编号
Private Const CONST_A01_FName As Long = 1 '产品名称
Private Const CONST_A01_FModel As Long = 2 '规格
Private Const CONST_A01_FBatchNo  As Long = 3 '产品批次
Private Const CONST_A01_FDate As Long = 4 '生产日期
Private Const CONST_A01_FBoxBarCode As Long = 5 ' 箱条码
Private Const CONST_A01_FBarCode As Long = 6 '盒条码
Public strGongChang As String



 
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

'---------------------------------------------------------------------------------------
' Procedure : SynBillA01
' DateTime  : 2013-1-25 00:05
' Author    :
' Purpose   : 生成外购入库单并审核
'---------------------------------------------------------------------------------------
'
Public Function SynPackage(ByVal sDsn As String, varDatas As KFO.Vector, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String, ByRef File() As Byte, ByRef strFileName As String, ByRef strUUID As String) As Boolean
    
 
    Dim i As Long
    Dim ssql As String
    Dim oconnect As Object
    Dim rs As ADODB.Recordset
    Dim Flag As Boolean
    
   
    Dim strItem As String
    
    Dim objTypeLib As Object
    
    Dim dctCheck As KFO.Dictionary
    Dim dctTemp As KFO.Dictionary
    Dim vctTemp As KFO.Vector
    Dim strAllSQL As New StringBuilder
    Dim strMsg As String
    Dim InterID As Long
    Dim strSQL As New StringBuilder
    Dim strFile As New StringBuilder
    Dim lngTaskID As Long
    Dim lngEntryID As Long
    Dim lngBillID As Long
'    Dim strFileName As String
    Dim objActivation As Object
    Dim strDSN2 As String '新帐套连接字符窜
    Dim BarCode As String
    Dim rsYearMonth As ADODB.Recordset
    Dim Count As Long
    Dim FileDate As String
    
'    Dim Conn As String
'    Dim soapclient As SoapClient30
'    Set soapclient = New SoapClient30

    Dim b As Boolean
    Dim stype As String
    Dim OutPut As String
    
    Set objTypeLib = Nothing
On Error GoTo HErr

        '使用GUID作为一次导入事务的标识
    Set objTypeLib = CreateObject("Scriptlet.TypeLib")
    strUUID = CStr(objTypeLib.Guid)
    strUUID = Mid(strUUID, 1, InStr(1, strUUID, "}"))

    SynPackage = False
    ssql = ""
    '将文件名上的日期导入到
    If InStr(strFileName, "__") = False Then
        sErr = sErr & "文件名格式有误"
        GoTo HErr
    End If
    FileDate = Left(Split(strFileName, "__")(1), 8)
'    strDSN2 = TransfersDsn2("114.215.188.89", "ReportServerTempDB", "sa", "W!ndows@2016", sDsn)
    strDSN2 = TransfersDsn("st20160407", sDsn)

    Set vctTemp = New KFO.Vector

    '========================导入条码=====================
    For i = 1 To varDatas.Size
         Flag = True
         Set dctTemp = New KFO.Dictionary
           
         If CheckProduct(varDatas(i)(CONST_A01_FBillNo), sDsn) = False Then
            sErr = sErr & vbCrLf & "产品条码： " & varDatas(i)(CONST_A01_FBillNo) & " 不存在"
            Flag = False
         End If


'
         If Flag = False Then
            GoTo Line1
         End If

'         ssql = "insert T_t_Package "
'         ssql = ssql & vbCrLf & "values('" & Trim(CStr(varDatas(i)(CONST_A01_FBillNo))) & "',"
'         ssql = ssql & "'" & Trim(CStr(varDatas(i)(CONST_A01_FName))) & "',"
'         ssql = ssql & "'" & Trim(CStr(varDatas(i)(CONST_A01_FModel))) & "',"
'         ssql = ssql & "'" & Trim(CStr(varDatas(i)(CONST_A01_FBatchNo))) & "',"
'         ssql = ssql & "'" & Format(varDatas(i)(CONST_A01_FDate), "yyyy-mm-dd") & "',"
'         ssql = ssql & "'" & Trim(Replace(CStr(varDatas(i)(CONST_A01_FBoxBarcode)), Chr(13), "")) & "',"
'         ssql = ssql & "'" & Trim(Replace(CStr(varDatas(i)(CONST_A01_FBarCode)), Chr(13), "")) & "')"
'
         strSQL.Append "insert T_t_Package "
         strSQL.Append "values('" & Trim(CStr(varDatas(i)(CONST_A01_FBillNo))) & "',"
         strSQL.Append "'" & Trim(CStr(varDatas(i)(CONST_A01_FName))) & "',"
         strSQL.Append "'" & Trim(CStr(varDatas(i)(CONST_A01_FModel))) & "',"
         strSQL.Append "'" & Trim(CStr(varDatas(i)(CONST_A01_FBatchNo))) & "',"
         strSQL.Append "'" & Format(varDatas(i)(CONST_A01_FDate), "yyyy-mm-dd") & "',"
         strSQL.Append "'" & Trim(Replace(CStr(varDatas(i)(CONST_A01_FBoxBarCode)), Chr(13), "")) & "',"
         strSQL.Append "'" & Trim(Replace(CStr(varDatas(i)(CONST_A01_FBarCode)), Chr(13), "")) & "',"
         strSQL.Append "'" & strUUID & "',"
         strSQL.Append "'" & Format(varDatas(i)(CONST_A01_FDate), "yyyy-mm") & "',"
         strSQL.Append "'" & FileDate & "')"
         
         dctTemp("sql") = strSQL.StringValue
         strSQL.Remove 1, strSQL.Length
         vctTemp.Add dctTemp
         Set dctTemp = Nothing
Line1:
    Next
    
    If sErr <> "" Then
        GoTo HErr
    End If
    
    strAllSQL.Append "set nocount on"
    
    For i = vctTemp.LBound To vctTemp.UBound
        strAllSQL.Append vbCrLf & vctTemp(i)("sql")
        If i Mod 100 = 0 Then
       '    Debug.Print strAllSQL
            Set oconnect = CreateObject("K3Connection.AppConnection")
            ExecSQL strAllSQL.StringValue, sDsn
            Set oconnect = Nothing
            strAllSQL.Remove 1, Len(strAllSQL.StringValue)
            strAllSQL.Append "set nocount on"
        End If
    Next


    If strAllSQL.StringValue <> "set nocount on" Then
      ' Debug.Print strAllSQL        Set oconnect = CreateObject("K3Connection.AppConnection")
        ExecSQL strAllSQL.StringValue, sDsn
        Set oconnect = Nothing
    End If
    strFileName = ""
   
'     '=================插入新数据库=====================
'        '获取新数据库的DSN
'        strDSN2 = TransfersDsn("st20160407", sDsn)
'
'
'        '查询本次导入的数据按到期日年月分组导入
'        ssql = "select isnull(FYearMonth,'') FYearMonth from T_t_Package where FUUID='" & strUUID & "' and isnull(FYearMonth,'') <>'' group by FYearMonth"
'        Set rsYearMonth = ExecSQL(ssql, sDsn)
'
'        For i = 1 To rsYearMonth.RecordCount
'
'        '根据到期日检查对应表在新数据库中是否存在,不存在则新建
'            ssql = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[T_t_Package_" & rsYearMonth.Fields("FYearMonth") & "]') AND type in (N'U'))"
'            ssql = ssql & vbCrLf & "CREATE TABLE [dbo].[T_t_Package_" & rsYearMonth.Fields("FYearMonth") & "]("
'            ssql = ssql & vbCrLf & "[FProductNumber] [nvarchar](255) NULL,"
'            ssql = ssql & vbCrLf & "[FProductName] [nvarchar](255) NULL,"
'            ssql = ssql & vbCrLf & "[FModel] [nvarchar](150) NULL,"
'            ssql = ssql & vbCrLf & "[FBatch] [nvarchar](50) NULL,"
'            ssql = ssql & vbCrLf & "[FDate] [nvarchar](50) NULL,"
'            ssql = ssql & vbCrLf & "[FBoxBarCode] [nvarchar](255) NULL,"
'            ssql = ssql & vbCrLf & "[FHeBarCode] [nvarchar](255) NULL,"
'            ssql = ssql & vbCrLf & "[FUUID] [nvarchar](255)NULL,"
'            ssql = ssql & vbCrLf & "[FYearMonth] [nvarchar](15)"
'            ssql = ssql & vbCrLf & "[FFileDate] [nvarchar](15)"
'            ssql = ssql & vbCrLf & ") ON [PRIMARY]"
'            ExecSQL ssql, strDSN2
'
'            '记录新建表名
'            TableName = TableName & "  T_t_Package_" & rsYearMonth.Fields("FYearMonth")
'
'            '将本次导入数据按到期年月分组导入对应新数据库表中
'            ssql = "insert into [T_t_Package_" & rsYearMonth.Fields("FYearMonth") & "] select * from [st20160414].dbo.T_t_Package where FUUID='" & strUUID & "' and FYearMonth='" & rsYearMonth.Fields("FYearMonth") & "'"
'            ExecSQL ssql, strDSN2
'
'            '检查重复盒条码
'            ssql = "select FHeBarCode from [t_t_package_" & rsYearMonth.Fields("FYearMonth") & "] where fhebarcode in (select FHeBarCode from [T_t_Package_" & rsYearMonth.Fields("FYearMonth") & "] group by FHeBarCode having COUNT(1)>=2) and fuuid='" & strUUID & "'"
'            Set rs = ExecSQL(ssql, strDSN2)
'            If rs.EOF = False Then
'                For Count = 1 To rs.RecordCount
'                    BarCode = BarCode & rs.Fields("FHeBarCode") & "  "
'                    rs.MoveNext
'                Next
'                sErr = "以下盒条码已存在 [" & Left(BarCode, Len(BarCode) - 1) & "]"
'                '将重复条码删除
'                ssql = "delete [t_t_package_" & rsYearMonth.Fields("FYearMonth") & "] where fhebarcode in (select FHeBarCode from [t_t_package_" & rsYearMonth.Fields("FYearMonth") & "] group by FHeBarCode having COUNT(1)>=2) and fuuid='" & strUUID & "'"
'                Set rs = ExecSQL(ssql, strDSN2)
'            End If
'
'        rsYearMonth.MoveNext
'        Next i
'    '==================================================
        
    
    '===================生成外购入库================
    '2017-10-16包装扫码信息不生成外购入库，仅同步数据
    '获取本次需要导入的数据
'    ssql = "select * from T_t_Package where fUUID='" & strUUID & "' order by FProductnumber,FBatch,FDate"
'    Set rs = ExecSQL(ssql, sDsn)
'    If rs.EOF Then
'        sErr = sErr & "没有符合条件的数据导入"
'        GoTo HErr
'    End If
'
'
'    If SynBillA01(sDsn, InterID, sBillNo, strMsg, rs) = False Then
'        sErr = sErr & strMsg
'        GetObjectContext.SetAbort
'        SynPackage = False
'        Exit Function
'    End If
    
    '===========生成二次激活文件==========
    lngTaskID = 1
    lngEntryID = 1
    lngBillID = 10000
    ssql = "select FBatch,FProductNumber,FHeBarCode,CONVERT(varchar(100), GETDATE(), 20)ExportTime from T_t_Package where FUUID='" & strUUID & "'"
    Set rs = ExecSQL(ssql, sDsn)
    If rs.EOF = False Then
        For i = 1 To rs.RecordCount
            If i = 1 Then '打印第一行Task
                strFile.Append "$" & lngTaskID & "," & lngBillID & "," & Replace(Replace(Replace(strUUID, "-", ""), "{", ""), "}", "") & ",0000001" & vbCrLf
            End If
            '打印Data部分
            strFile.Append "!" & lngTaskID & "," & lngEntryID & "," & rs.Fields("FBatch") & "," & rs.Fields("FProductNumber") & "," & rs.Fields("FHeBarCode") & "," & rs.Fields("ExportTime") & vbCrLf
            lngEntryID = lngEntryID + 1
        rs.MoveNext
        Next
    End If

    strFileName = App.Path & "\二次激活_" & Format(Now, "yyyyMMddHHmmss") & ".txt"

    Open strFileName For Output As #123
    Print #123, strFile.StringValue
    Close #123

    File = StrConv(strFile.StringValue, vbFromUnicode)
'    soapclient.MSSoapInit "http://dt.yesno.com.cn/DataTransferWebService/DataTransferWebService.asmx?WSDL"
'    soapclient.MSSoapInit2 "http://dt.yesno.com.cn/DataTransferWebService/DataTransferWebService.asmx?WSDL", "", "", "", ""
'    b = soapclient.UploadBytesFile("9457", "0000001", File, strFileName, 1, "S0001", "S0001@2016", "1.1.0.1", stype)
    '=====================================
    
        
    '导入完成后将原数据库的数据删除
'    ExecSQL "delete t_t_package where fuuid='" & strUUID & "'", sDsn
    
   
    SynPackage = True
    GetObjectContext.SetComplete
    Exit Function
HErr:
    SynPackage = False
    GetObjectContext.SetAbort
    sErr = "Import BZSM error：" & sErr & Err.Description
End Function

Public Function ExecSQL(ByVal ssql As String, ByVal dsn As String) As ADODB.Recordset
    Dim OBJ As Object
    Dim rs As ADODB.Recordset

    Set OBJ = CreateObject("BillDataAccess.GetData")
    Set rs = OBJ.ExecuteSQL(dsn, ssql)
    Set OBJ = Nothing
    Set ExecSQL = rs
    Set rs = Nothing
End Function


Private Function CheckProduct(ByVal ProductNumber As String, dsn As String) As Boolean
Dim rs As ADODB.Recordset

Set rs = ExecSQL("select 1 from t_ICItem where FBarCode='" & ProductNumber & "'", dsn)
If rs.RecordCount = 0 Then
    CheckProduct = False
    Exit Function
End If
CheckProduct = True
End Function

Private Function CheckBarCode(ByVal FHeBarCode As String, dsn As String) As Boolean
Dim rs As ADODB.Recordset

Set rs = ExecSQL("select 1 from t_t_Package where replace(FHeBarCode,CHAR(13),'')='" & FHeBarCode & "'", dsn)
If rs.RecordCount = 0 Then
    CheckBarCode = True
    Exit Function
End If
CheckBarCode = False
End Function


'---------------------------------------------------------------------------------------
' Procedure : SynBillA01
' DateTime  : 2013-1-25 00:05
' Author    :
' Purpose   : 生成外购入库单并审核
'---------------------------------------------------------------------------------------
'
Public Function SynBillA01(ByVal sDsn As String, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String, rsBill As ADODB.Recordset) As Boolean
    Dim CNN As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim clsBill As AutoOldBill
    Dim bRet As Boolean
    Dim lRow As Long
    Dim i As Long
    Dim dblDiscountAmount As Double '折扣额
    Dim dblMarkDownAmt As Double ''汇总折扣代码为 PRWN 的金额
    Dim dblAmt As Double '金额
    Dim dblDiscount As Double
    
    Dim lFItemID As Long
    Dim lUnitID As Long
    Dim dRate As Double
    
    Dim lSPGroupID As Long
    Dim blnIsSPGrop As Boolean
    Dim blnCheck As Boolean
    Dim sSourceBillNo As String
    Dim lROB As Long
    Dim blnEffective As Boolean ''是否存在有效物料
    Dim isUpdateStore As Boolean
    
    Dim strMarkDown As String
    
    Dim strBatch As String '保存批号用于比较,同一物料对批号保质期一致的数量汇总
    Dim strPeriodDate As String
    Dim strProduct As String
    
    Dim lngQty As Long '数量
    Dim lngPeriod As Long '保质期
    Dim lngStock As Long '仓库
    
    Dim lPOOrderID As Long
    Dim lStockID As Long
    Dim strSQL As String
    Dim Row As Long
    Dim rsField As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    Dim str() As String
    
On Error GoTo SynBillA01_Error

    InitDataEnv sDsn
    Set CNN = datasource.Connection
    
    Set clsBill = New AutoOldBill
    
    lRow = 1
    blnCheck = False
    
    blnEffective = False
'    GetObjectContext.SetAbort

        '供应商: 默认取值G017公司
        '日期: 取单据生成日期
        '编号: 根据编码规则自动生成
        '收料仓库: 取商品信息中的默认仓库 (C001?C002?C006?C014)
        '仓位：默认为*
        
        strSQL = "select FSQLColumnName from t_ItemPropDesc where FItemClassID=4 and FName like '包装关联工厂'"
        Set rsField = CNN.Execute(strSQL)
        If rsField.RecordCount = 0 Then
            sErr = "Can not find the field [工厂]"
            Exit Function ' rsField.Fields("FSQLColumnName") = ""
        Else
            strGongChang = rsField.Fields("FSQLColumnName")
        End If
        
    lROB = 1
    
    With clsBill
        .sDsn = sDsn
        
        bRet = .NewBill(1)

        If bRet = True Then
            For i = 1 To rsBill.RecordCount
                
                If i = 1 Then '表头
    

                    .SetHeadValue "FSettleDate", Format(Now, "yyyy-mm-dd") '付款日期
                    .SetHeadValue "FDate", Format(Now, "yyyy-mm-dd") '日期
                    
                    
                    lFItemID = GetItemID(CNN, "G017", 8)
                    If lFItemID = 0 Then
                        sErr = "供应商[G017] 不存在"
                        Exit Function
                    End If
                    
                    .SetHeadValue "FSupplyID", lFItemID '供应商默认取采购订单供应商
                    .SetHeadValue "FPOMode", 36680 '采购模式:  默认为普通采购
                    .SetHeadValue "FPOStyle", 252  '采购方式: 默认为赊购
                    
'                    '仓库
'                    lStockID = GetItemID(CNN, "C004", 5)
'                    If lStockID = 0 Then
'                        sErr = "仓库[C004] 不存在"
'                        Exit Function
'                    End If
                    

                    lFItemID = GetItemID(CNN, "037", 3)
                    If lFItemID = 0 Then
                        sErr = "Not found the code for the [037]  staff"
                        Exit Function
                    End If
                    .SetHeadValue "FEmpID", lFItemID
                    .SetHeadValue "FFManagerID", lFItemID
                    .SetHeadValue "FManagerID", lFItemID
                    .SetHeadValue "FSManagerID", lFItemID
                    
                    
                    lngQty = 1
                End If
                
                '表体赋值
                      
                .AddRow lRow
                
                If strProduct = rsBill.Fields("FProductNumber") And strBatch = rsBill.Fields("FBatch") And strPeriodDate = rsBill.Fields("FDate") Then
                    lngQty = lngQty + 1
                    .SetEntryValue "Fauxqty", lngQty, Row              '数量
                    GoTo NewRow
                Else
                    lngQty = 1
                End If
                
                lFItemID = GetICItemID(CNN, rsBill.Fields("FProductNumber"), lUnitID, lngStock, lngPeriod, strGongChang)
                If lFItemID = 0 Then
                    sErr = "Generate Purchase Receive error: BarCode [" & rsBill.Fields("FProductNumber") & "] does not exit"
                    Exit Function
                End If
                
                If lngStock = 0 Then
                    sErr = "Generate Purchase Receive error: BarCode [" & rsBill.Fields("FProductNumber") & "] Default Stock does not set"
                    Exit Function
                End If

                Row = lRow
                .SetEntryValue "FItemID", lFItemID, lRow            '物料
                strProduct = rsBill.Fields("FProductNumber")
                
                .SetEntryValue "FUnitID", lUnitID, lRow            '单位
                .SetEntryValue "Fauxqty", lngQty, lRow              '数量
                
                .SetEntryValue "FDCStockID", lngStock, lRow              ''仓库
                
                strSQL = "select top 1 t2.FSPID from t_Stock t1 inner join t_StockPlace t2 on t1.FSPGroupID=t2.FSPGroupID where t1.FItemID=" & lngStock & " and t1.FSPGroupID<>0 order by t2.FSPID "
                Set rs = CNN.Execute(strSQL)
                
                If rs.RecordCount = 0 Then
                    .SetEntryValue "FDCSPID", 0, lRow              ''仓位
                Else
                    .SetEntryValue "FDCSPID", rs.Fields("FSPID"), lRow              ''仓位
                End If
                
                
                
                .SetEntryValue "FBatchNo", rsBill.Fields("FBatch"), lRow              ''批号
                strBatch = rsBill.Fields("FBatch")
                
                .SetEntryValue "FKFDate", DateAdd("d", -lngPeriod, rsBill.Fields("FDate")), lRow              ''生产日期
                
                
                .SetEntryValue "FKFPeriod", lngPeriod, lRow              ''保质期
                .SetEntryValue "FPeriodDate", rsBill.Fields("FDate"), lRow          ''到期日
                strPeriodDate = rsBill.Fields("FDate")
               
                lRow = lRow + 1
NewRow:
            rsBill.MoveNext
            Next
            
            bRet = .SaveBill(lInterID, sBillNo, sErr)
'            If bRet = True Then
'                checkBillData sDsn, lInterID, 1, 0, sErr, 16394 '审核单据
'            End If
        End If
    End With
    
    Set clsBill = Nothing
    Set CNN = Nothing
    SynBillA01 = True
    GetObjectContext.SetComplete

   Exit Function

SynBillA01_Error:
    SynBillA01 = False
'    GetObjectContext.SetAbort
    Set clsBill = Nothing
    Set CNN = Nothing
    sErr = "Generate Puchase Receive error：" & sErr & Err.Description
    
End Function



