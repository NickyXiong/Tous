Attribute VB_Name = "Globals"
' Kingdee Enterprise Business Objects
' Copyright (C) 1995-1998 Kingdee Corporation
' All rights reserved

Option Explicit
Private Declare Function VariantChangeType Lib "oleaut32.dll" ( _
    pvarDest As Variant, pvarSrc As Variant, ByVal wFlags As Integer, ByVal vt As Integer) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function ReadTcp Lib "kdappsvr.dll" (cComputer As String, cVal As String) As Long
'////////////////////////////////////////////////////////////////////////
'// Security functions
Public Type T_ITEMDETAILS
    ItemClassID As Long
    ItemID As Long
End Type
Public g_objResLoader As Object '多语言函数加载对象
'英语添加空格的方式
Enum EnAppendBlank
    EnAppendBlank_NULL
    EnAppendBlank_PREV
    EnAppendBlank_POST
    EnAppendBlank_BOTH
End Enum

Public Function GetClientUserID(ByVal datasource As CDataSource) As Integer
    Dim sec As Object   'New EBSBase.SecurityInfo
    Set sec = GetObjectContext.CreateInstance("EBSBase.SecurityInfo")
    GetClientUserID = sec.GetCurrentUserID(datasource.ParseObject.PropsString)
End Function

Public Sub AccessCheck(ByVal datasource As CDataSource, ByVal ObjectType As GLObjectTypeEnum, _
    ByVal ObjectID As Long, ByVal DesiredAccess As Long, ByVal Source As String)

    Dim sec As Object    'EBSBase.SecurityInfo
    Set sec = GetObjectContext.CreateInstance("EBSBase.SecurityInfo")
    sec.AccessCheck datasource.ParseObject.PropsString, ObjectType, ObjectID, DesiredAccess, True, Source
End Sub

Public Function GetSystemProfile(ByVal datasource As CDataSource, _
    ByVal Key As String, _
    Optional ByVal Category As String = "GL", _
    Optional ByVal Default As Variant) As Variant
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = datasource.Connection
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select FValue From t_SystemProfile Where FCategory = ? And FKey = ?"
    cmd.Parameters.Append cmd.CreateParameter( _
        Type:=adVarChar, Size:=30, Value:=Category)
    cmd.Parameters.Append cmd.CreateParameter( _
        Type:=adVarChar, Size:=30, Value:=Key)
    With cmd.Execute()
        If Not .EOF Then
            GetSystemProfile = .Fields(FValue).Value
        Else
            If IsMissing(Default) Then
                GetSystemProfile = ""
            Else
                GetSystemProfile = Default
            End If
        End If
    End With
    Set cmd = Nothing
End Function

'////////////////////////////////////////////////////////////////////////
'// Check functions, these functions may throw exceptions

Public Sub CheckVariant(ByVal VarName As String, ByRef Value As Variant, _
    ByVal vt As VbVarType, Optional ByVal Size As Long = 0, _
    Optional ByVal Required As Boolean = False)

    ' Handle null values
    If IsNull(Value) Or IsEmpty(Value) Then
        If Required Then
            Err.Raise EBS_E_MissingRequiredData, "EBSGL.Globals.CheckVariant", "Variable ='" & VarName & "'"
        End If
        Value = Null
        Exit Sub
    End If
    
    ' Attempt to convert to required type
    If VarType(Value) <> vt Then
        Dim hr As Long
        hr = VariantChangeType(Value, Value, 0, vt)
        Select Case hr
        Case &H80020005 ' DISP_E_TYPEMISMATCH
            hr = EBS_E_TypeMismatch
        Case &H8002000A ' DISP_E_OVERFLOW
            hr = EBS_E_DataOverflow
        Case &H80020008 ' DISP_E_BADVARTYPE
            Debug.Assert False
        End Select
        If hr <> 0 Then Err.Raise hr, "EBSGL.Globals.CheckVariant", "Variable ='" & VarName & "'"
    End If
    
    Dim nLen As Long
    
    If VarType(Value) = vbString Then
        ' Handle string values
        nLen = LenA(Value)
        If nLen = 0 Then
            ' treat empty string as Null
            If Required Then
                Err.Raise EBS_E_MissingRequiredData, "EBSGL.Globals.CheckVariant", , "Variable ='" & VarName & "'"
            End If
            Value = Null
        ElseIf Size > 0 And nLen > Size Then
            Err.Raise EBS_E_DataTooLong, "EBSGL.Globals.CheckVariant", "Variable ='" & VarName & "' Value = " & Value
        End If
    ElseIf VarType(Value) = (vbArray + vbByte) Then
        ' Handle binary data
        nLen = UBound(Value) - LBound(Value) + 1
        If nLen = 0 Then
            ' treat zero length array as Null
            If Required Then
                Err.Raise EBS_E_MissingRequiredData, "EBSGL.Globals.CheckVariant", "Variable ='" & VarName & "'"
            End If
            Value = Null
        ElseIf Size > 0 And nLen > Size Then
            Err.Raise EBS_E_DataTooLong, "EBSGL.Globals.CheckVariant", "Variable ='" & VarName & "' Value = " & Value
        End If
    End If
End Sub

Public Function VarTypeFromDbType(ByVal DbType As ADODB.DataTypeEnum) As Integer
    Select Case DbType
    Case adBoolean
        VarTypeFromDbType = vbBoolean
    Case adTinyInt
        VarTypeFromDbType = vbByte
    Case adSmallInt
        VarTypeFromDbType = vbInteger
    Case adInteger
        VarTypeFromDbType = vbLong
    Case adBigInt
        VarTypeFromDbType = vbString
    Case adSingle
        VarTypeFromDbType = vbSingle
    Case adDouble
        VarTypeFromDbType = vbDouble
    Case adNumeric, adDecimal
        VarTypeFromDbType = vbString
    Case adCurrency
        VarTypeFromDbType = vbCurrency
    Case adDate, adDBDate, adDBTime, adDBTimeStamp
        VarTypeFromDbType = vbDate
    Case adChar, adVarChar, adLongVarChar
        VarTypeFromDbType = vbString
    Case adBinary, adVarBinary, adLongVarBinary
        VarTypeFromDbType = vbArray + vbByte
    Case Else
        Debug.Assert False
        VarTypeFromDbType = vbEmpty
    End Select
End Function

Public Function IsValidID(ByVal ID As String, _
        Optional ByVal MaxLength As Integer = 0, _
        Optional ByVal AllowUnderscore As Boolean = True) As Boolean
        
    Const ASC_A = 65, ASC_Z = 90
    Const ASC_0 = 48, ASC_9 = 57
    Const ASC_UNDERSCORE = 95
    If ID = "0" Then Exit Function
    Dim nLen As Integer
    nLen = Len(ID)
    If MaxLength > 0 And nLen > MaxLength Then
        Exit Function
    End If

    Dim i As Integer
    For i = 1 To nLen
        Dim ch As Integer
        ch = Asc(Mid(ID, i))
        If Not ((ch >= ASC_A And ch <= ASC_Z) Or _
                (ch >= ASC_0 And ch <= ASC_9) Or _
                (AllowUnderscore And ch = ASC_UNDERSCORE)) Then
            Exit Function
        End If
    Next

    IsValidID = True
End Function



'////////////////////////////////////////////////////////////////////////
'// Helper functions

Public Function PackageRecord(rs As ADODB.Recordset) As Object
    Debug.Assert Not rs.EOF And Not rs.BOF
    Dim dict As New KFO.Dictionary
    Dim fld As ADODB.Field
    For Each fld In rs.Fields
        dict(fld.Name) = fld.Value
    Next
    Set PackageRecord = dict
End Function

Public Function PackageRecordset(rs As ADODB.Recordset) As Object
    Dim vec As New KFO.Vector
    Do Until rs.EOF
        vec.Add PackageRecord(rs)
        rs.MoveNext
    Loop
    Set PackageRecordset = vec
End Function

Public Sub ResetParams(cmd As ADODB.Command)
    With cmd.Parameters
        While .Count <> 0
            .Delete 0
        Wend
    End With
End Sub

Public Sub GrowSize(ByVal Count As Long, ByRef MaxCount As Long)
    If Count > MaxCount Then
        Dim nGrowBy As Long
        ' heuristically determine growth (this avoids heap fragmentation
        ' in many situations)
        nGrowBy = Count \ 8
        If nGrowBy < 4 Then
            nGrowBy = 4
        ElseIf nGrowBy > 1024 Then
            nGrowBy = 1024
        End If
        
        Dim nNewMax As Long
        If Count < MaxCount + nGrowBy Then
            nNewMax = MaxCount + nGrowBy ' granularity
        Else
            nNewMax = Count ' no slush
        End If
        
        Debug.Assert nNewMax >= MaxCount ' no warp around
        MaxCount = nNewMax
    End If
End Sub

' Check Account Group
Public Function CheckAccountGroup(ByVal datasource As CDataSource, ByVal GroupID As Long) As Boolean
    Dim cn As ADODB.Connection
    Set cn = datasource.Connection
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("Select FGroupID From t_AcctGroup Where FGroupID=" & GroupID)
    CheckAccountGroup = Not rs.EOF
End Function
'////////////////////////////////////////////////////////////////////////
'// Check to see weither a specified object already in use
Public Function ObjectInUsed(ByVal datasource As CDataSource, _
            ByVal ObjectType As GLObjectTypeEnum, _
            ByVal ObjectID As Long) As Boolean
    If Not datasource.ProcSupported("sp_ObjectInUsed") Then
        Err.Raise EBS_E_ObjectNotFound, , "sp_ObjectInUsed Not Found"
    End If
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = datasource.Connection
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "sp_ObjectInUsed"
    cmd.Parameters.Append cmd.CreateParameter( _
                Type:=adInteger, Direction:=adParamReturnValue)
    cmd.Parameters.Append cmd.CreateParameter( _
                Type:=adInteger, Value:=ObjectType)
    cmd.Parameters.Append cmd.CreateParameter( _
                Type:=adInteger, Value:=ObjectID)
    cmd.Execute
    ObjectInUsed = cmd.Parameters(0).Value
End Function

'////////////////////////////////////////////////////////////////////////
'// Voucher object global functions
Private Function PeriodDate(ByVal Year As Integer, ByVal PeriodDates As String, ByVal PeriodCount As Long, ByVal Period As Long) As Date
    Year = (Period - 1) \ PeriodCount + Year
    Period = (Period - 1) Mod PeriodCount + 1
    PeriodDate = DateSerial(Year, CInt(Mid(PeriodDates, Period * 4 - 3, 2)), CInt(Mid(PeriodDates, Period * 4 - 1, 2)))
End Function
Private Function YearDiff(ByVal dt1 As Date, ByVal dt2 As Date) As Integer
    Dim i As Integer, m As Integer, d As Integer
    m = Month(dt1): d = Day(dt1)
    i = 0
    If dt1 > dt2 Then
        While dt1 > dt2
            dt1 = DateSerial(Year(dt1) - 1, m, d)
            i = i - 1
        Wend
    ElseIf dt1 < dt2 Then
        While dt1 < dt2
            dt1 = DateSerial(Year(dt1) + 1, m, d)
            i = i + 1
        Wend
        If dt1 > dt2 Then i = i - 1
    End If
    YearDiff = i
End Function
Public Function PeriodFromDate(ByVal datasource As CDataSource, ByVal dt As Date, nYear As Long) As Long
'    If DataSource.ProcSupported("sp_PeriodFromDate") Then
'        PeriodFromDate = sp_PeriodFromDate(dt, nYear)
'        Exit Function
'    End If
    
    Dim CurrYear As Integer
    Dim PeriodCount As Integer
    Dim PeriodDates As String
    Dim PeriodYearDiff As Long
    
    Dim YearStartDate As Date
    Dim Period As Long
    
    CurrYear = CInt(GetSystemProfile(datasource, GLCurrentYear))
    PeriodCount = CInt(GetSystemProfile(datasource, GLPeriodCount))
    PeriodDates = GetSystemProfile(datasource, GLPeriodDates)
    PeriodYearDiff = Val(GetSystemProfile(datasource, GLYearDifference))
    YearStartDate = PeriodDate(CurrYear, PeriodDates, PeriodCount, 1)
    nYear = CurrYear + YearDiff(YearStartDate, dt)
    Dim dt2 As Date
    Period = 1
    dt2 = PeriodDate(nYear, PeriodDates, PeriodCount, Period)
    While dt > dt2
        Period = Period + 1
        dt2 = PeriodDate(nYear, PeriodDates, PeriodCount, Period)
    Wend
    If dt < dt2 Then
        Period = Period - 1
    End If
    nYear = nYear + PeriodYearDiff
    PeriodFromDate = Period
End Function

Private Function sp_PeriodFromDate(ByVal datasource As CDataSource, ByVal dt As Date, nYear As Long) As Long
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = datasource.Connection
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "sp_PeriodFromDate"
    cmd.Parameters.Append cmd.CreateParameter( _
        Type:=adDBTimeStamp, Value:=dt)
    cmd.Parameters.Append cmd.CreateParameter( _
        Type:=adInteger, Direction:=adParamOutput)
    cmd.Parameters.Append cmd.CreateParameter( _
        Type:=adInteger, Direction:=adParamOutput)
        
    cmd.Execute
    nYear = cmd(1)
    sp_PeriodFromDate = cmd(2)
End Function

Public Function IsFraction(ByVal n As Double) As Boolean
    IsFraction = (n <> Val(str(n)))
End Function

Public Function SeparateString(ByVal PathString As String, _
                        Optional ByVal SepChar = vbTab) As Collection
    Dim mCol As New Collection
    Dim i As Long
    PathString = Trim(PathString)
    i = InStr(1, PathString, SepChar)
    While i > 0
        mCol.Add Trim(Left(PathString, i - 1))
        PathString = Trim(Mid(PathString, i + 1))
        i = InStr(1, PathString, SepChar)
    Wend
    If PathString <> "" Then mCol.Add PathString
    Set SeparateString = mCol
End Function
Public Function IsDetail(ByVal datasource As CDataSource, ByVal m_AccountID As Long) As Boolean
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    
    SQL = "SELECT Fdetail from t_Account " & _
            "Where  FAccountID = " & m_AccountID
    With rs
        .CursorLocation = adUseClient
        .Open SQL, datasource.Connection
        IsDetail = rs!FDetail
        rs.Close
    End With

End Function



'////////////////////////////////////////////////////////////////////////
Public Function GetSubNumber(ByVal Number As String, Optional ParentNumber As String) As String
    Dim i As Long
    i = InStrRev(Number, ".")
    If i <> 0 Then
        GetSubNumber = Mid(Number, i + 1)
        ParentNumber = Left$(Number, i - 1)
    Else
        GetSubNumber = Number
        ParentNumber = ""
    End If
End Function
'Public Function CreateDefaultItem( _
'                Optional ByVal ItemClassID As Long, _
'                Optional ByVal ItemID As Long) As Long
'
'    Debug.Assert ItemClassID <> 0 Or ItemID <> 0
'    Dim cn As ADODB.Connection
'    Dim strSQL As String
'    Set cn = DataSource.Connection
'    cn.CursorLocation = adUseClient
'    Dim Number As String
'    Dim bDetail As Boolean
'    Dim nLevel As Integer
'    Dim ParentID As Long
'    Dim cmd As ADODB.Command
'    Dim rs As ADODB.Recordset
'    Dim cb As New CommandBuilder
'    If ItemClassID = 0 And ItemID = 0 Then Exit Function
'    If ItemID <> 0 Then
'        Set rs = cn.Execute("Select FItemClassID,FNumber,FDetail From t_Item Where FItemID=" & ItemID)
'        If rs.EOF Then
'            Err.Raise EBS_E_ObjectNotFound, "EBSGL.Globals.CreateDefaultItem", "ItemID=" & ItemID
'        End If
'        Number = rs!FNumber & ".0"
'        bDetail = rs!FDetail
'        nLevel = rs!FLevel + 1
'        ItemClassID = rs!FItemClassID
'        ParentID = ItemID
'        rs.Close
'        strSQL = "Select FItemID,FNumber From t_Item Where FParentID=" & ItemID & " And FDefault <> 0"
'        Set rs = cn.Execute(strSQL)
'        If Not rs.EOF Then
'            Err.Raise EBS_E_ObjectExists, "EBSGL.Globals.CreateDefaultItem", "Default Item = " & rs!FNumber
'        End If
'        rs.Close
'    Else
'        Set rs = cn.Execute("Select FItemClassID From t_ItemClass Where FItemClassID=" & ItemClassID)
'        If rs.EOF Then
'            Err.Raise EBS_E_ObjectNotFound, "EBSGL.Globals.CreateDefaultItem", "ItemClassID=" & ItemClassID
'        End If
'        rs.Close
'        Set rs = cn.Execute("Select FItemID,FNumber From t_Item Where FParentID=0 And FItemClassID=" & ItemClassID & " And FDefault<>0")
'        If Not rs.EOF Then
'            Err.Raise EBS_E_ObjectExists, "EBSGL.Globals.CreateDefaultItem", "Default Item = " & rs!FNumber
'        End If
'        rs.Close
'        Number = "0"
'        nLevel = 0
'        ParentID = 0
'    End If
'    cb.Create t_Item
'    cb.DescribField FItemClassID, adInteger, , True, ItemClassID
'    cb.DescribField FNumber, adVarChar, 80, True, Number & ".0"
'    cb.DescribField FLevel, adSmallInt, , True, nLevel
'    cb.DescribField FParentID, adInteger, , True, ParentID
'    cb.DescribField FDetail, adBoolean, , True, True
'
'End Function
Public Sub CacheProfile( _
    ByVal Category As String, _
    ByVal Key As String, _
    Value As Variant)
    
    ' Use lower case string to compare
    Category = LCase$(Category)
    Key = LCase$(Key)
    
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProp As SharedProperty
    Dim bExists As Boolean
    
    Set spmMgr = CreateObject("MTxSpm.SharedPropertyGroupManager.1")
    Set spmGroup = spmMgr.CreatePropertyGroup("CachedProfiles", _
        LockSetGet, Process, bExists)
    Set spmProp = spmGroup.CreateProperty(Category, bExists)
    
    Dim dict As KFO.Dictionary
    If Not bExists Then
        Set dict = New KFO.Dictionary
        spmProp.Value = dict
    Else
        Set dict = spmProp.Value
    End If
    
    dict(Key) = Value
End Sub


Public Function PeriodToDate(ByVal datasource As CDataSource, ByVal nYear As Long, ByVal Period As Long, _
                Optional ByVal PeriodCount As Integer = 0, _
                Optional ByVal PeriodDates As String = vbNullString, _
                Optional ByVal YearDifference As Integer = 0) As Date
    If PeriodCount = 0 Then
        PeriodCount = Val(GetSystemProfile(datasource, GLPeriodCount))
    End If
    If PeriodDates = vbNullString Then
        PeriodDates = GetSystemProfile(datasource, GLPeriodDates)
        YearDifference = Val(GetSystemProfile(datasource, GLYearDifference))
    End If
    PeriodToDate = PeriodDate(nYear - YearDifference, PeriodDates, PeriodCount, Period)
End Function

Public Function RoundCurrency(ByVal n As Currency, ByVal Scalex As Integer) As Currency
    RoundCurrency = Round(n, Scalex)
End Function
Public Function EnCodeSqlString(ByVal s As String) As String
    Dim i As Long
    Dim sTemp As String
    sTemp = ""
    i = InStr(1, s, "'")
    While i > 0
        sTemp = sTemp & Left$(s, i) & "'"
        s = Mid$(s, i + 1)
        i = InStr(1, s, "'")
    Wend
    If s <> vbNullString Then
        sTemp = sTemp & s
    End If
    EnCodeSqlString = sTemp
End Function
Public Sub CheckIsInitClosed(ByVal datasource As CDataSource)
    If VBA.Val(GetSystemProfile(datasource, GLInitClosed)) = 0 Then
        Err.Raise EBSGL_E_InitializeNotFinished
    End If
End Sub
Public Sub InsertVector(v As KFO.Vector, ByVal Index As Long, OBJ As Object)
    Dim i As Long
    v.Add v(v.UBound)
    For i = v.UBound - 1 To Index + 1 Step -1
        Set v(i) = v(i - 1)
    Next i
    Set v(Index) = OBJ
End Sub
'取得任一多核算组合的DetailID
'bAddNew = True 时,当前数据库中如果无这一组合的DetailID,自动增加一条记录
'GetItemDetail1中的Dictionary用法:
'        Dim Details As New KFO.Dictionary
'        ItemClassID = 1: ItemID = 1002
'        Details(ItemClassID) = ItemID
'        ItemClassID = 2: ItemID = 1003
'        Details(ItemClassID) = 1003
'        Dim DetailID As Long
'        DetailID = GetItemDetail1(Details)
        
Public Function GetItemDetail(ByVal datasource As CDataSource, _
                Details() As T_ITEMDETAILS, _
                ByVal DetailCount As Long, _
                Optional ByVal bAddNew As Boolean = True, _
                Optional cn As ADODB.Connection) As Long
    GetItemDetail = 0
    Dim i As Long
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    If cn Is Nothing Then
        Set cn = datasource.Connection
    End If
    Dim Filter As String
    Dim Count As Long
    Count = 0
    Dim ItemClassID As Long, ItemID As Long
    Dim flds As String, vals As String
    For i = LBound(Details) To LBound(Details) + DetailCount - 1
        ItemClassID = Details(i).ItemClassID
        ItemID = Details(i).ItemID
        If ItemID <> 0 Then
            Filter = Filter & " And F" & ItemClassID & "=" & ItemID
            flds = flds & ",F" & ItemClassID
            vals = vals & "," & ItemID
            Count = Count + 1
        End If
    Next i
    Debug.Assert Count > 0
    strSQL = "Select FDetailID From t_ItemDetail " & _
                "Where FDetailCount = " & Count & Filter
    Set rs = cn.Execute(strSQL)
    If rs.RecordCount > 0 Then
        GetItemDetail = rs.Fields(0).Value
        Exit Function
    End If
    If bAddNew = False Then
        Exit Function
    End If
    Dim lrows As Long
    Dim DetailID As Long
    strSQL = "Insert Into t_ItemDetail(FDetailCount" & flds & _
            ") values(" & Count & vals & ")"
    cn.Execute strSQL, lrows
    Debug.Assert lrows = 1
    Set rs = cn.Execute("Select MAX(FDetailID) From t_ItemDetail")
    DetailID = rs.Fields(0).Value
    For i = LBound(Details) To UBound(Details)
        ItemClassID = Details(i).ItemClassID
        ItemID = Details(i).ItemID
        If ItemID <> 0 Then
            strSQL = "Insert Into t_ItemDetailV" & _
                    "(FDetailID,FItemClassID,FItemID)" & _
                    "Values(" & DetailID & "," & _
                    ItemClassID & "," & ItemID & ")"
            cn.Execute strSQL, lrows
            Debug.Assert lrows = 1
        End If
    Next i
    GetItemDetail = DetailID
End Function

Public Function GetItemDetail1(ByVal datasource As CDataSource, _
                ByVal Details As KFO.Dictionary, _
                Optional ByVal bAddNew As Boolean = True, _
                Optional cn As ADODB.Connection) As Long
    GetItemDetail1 = 0
    If Details Is Nothing Then Exit Function
    If Details.Count = 0 Then Exit Function
    Dim ItemClassID As Long, ItemID As Long
    Dim i As Long
    Dim Count As Long
    Count = 0
    Dim items() As T_ITEMDETAILS
    For i = 1 To Details.Count
        ItemClassID = CLng(Details.Name(i))
        ItemID = CLng(Details.Value(Details.Name(i)))
        If ItemID <> 0 Then
            Count = Count + 1
            ReDim Preserve items(1 To Count)
            items(Count).ItemClassID = ItemClassID
            items(Count).ItemID = ItemID
        End If
    Next i
    GetItemDetail1 = GetItemDetail(datasource, items, Count, bAddNew, cn)
End Function

Public Function CloneRecordSet(ByVal rs As ADODB.Recordset) As ADODB.Recordset
    On Error Resume Next
    Dim f As ADODB.Field
    Dim rsNew As ADODB.Recordset
    
    Set rsNew = New ADODB.Recordset
    rsNew.CursorLocation = adUseClient
    With rsNew.Fields
        For Each f In rs.Fields
            Select Case f.Type
            Case ADODB.adDouble, ADODB.adNumeric
                .Append f.Name, ADODB.adDouble, f.DefinedSize, f.Attributes
            Case Else
                .Append f.Name, f.Type, f.DefinedSize, f.Attributes
            End Select
            
        Next
    End With
    rsNew.Open
    rs.Filter = ""
    rs.MoveFirst
    With rsNew
        'For i = 1 To rs.RecordCount
        Do While Not rs.EOF
            .AddNew
            For Each f In rs.Fields
                '.Fields(f.Name).Value = CNulls(f.Value, Empty)
                .Fields(f.Name).Value = f.Value
            Next
            .UpdateBatch
            rs.MoveNext
        Loop
        'Next i
    End With
    rsNew.MoveFirst
    Set CloneRecordSet = rsNew
    Set rsNew = Nothing
    Exit Function
errhandle:
    MsgBox Err.Description
End Function

'/* MACH added by hefan2002.03.14
'/* Function:四舍五入
'/* Input: 目标数据，保留的小数位
Public Function MyRound(ByVal pNumber As Variant, ByVal pNumDigitsAfterDecimal As Long) As Variant
On Error GoTo HERROR
    Dim lTmp As Variant
    Dim i As Long
    Dim sTmp As String
    Dim lHeadLen As Long
    Dim iGetRidofE As Integer
    
    If Not (IsNumeric(pNumber)) Then
        MyRound = 0
        Exit Function
    End If
    
    If pNumDigitsAfterDecimal >= 0 Then
           
        lTmp = 10 ^ pNumDigitsAfterDecimal
        
        pNumber = CDec(pNumber)
        
        pNumber = CDec(pNumber + IIf(pNumber < 0, -1, 1) * 0.5 / lTmp)
        
        sTmp = CDec(pNumber)
        
        lHeadLen = InStr(sTmp, ".")
        If lHeadLen = 0 Then lHeadLen = Len(sTmp)
        
        sTmp = VBA.Left(sTmp, lHeadLen + pNumDigitsAfterDecimal)
        
        If IsNumeric(sTmp) Then
            MyRound = CDec(sTmp)
        Else
            MyRound = 0
        End If
    Else
        MyRound = pNumber
    End If
    
    Exit Function
HERROR:
    Err.Clear
    MyRound = pNumber
End Function


Public Function GetLocalMachineName() As String
    Dim sBuffer As String * 128
    Dim lSize As Long
        
    lSize = 128
    If GetComputerName(sBuffer, lSize) > 0 Then
        GetLocalMachineName = VBA.UCase(VBA.Left(sBuffer, InStr(1, sBuffer, Chr(0)) - 1))
      Else
        GetLocalMachineName = ""
    End If
End Function

Public Function getLocalIPAddress() As String
    Dim sBuffer As String * 64
    If ReadTcp("", sBuffer) = 1 Then
        getLocalIPAddress = VBA.Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        getLocalIPAddress = ""
    End If
End Function
'***********************************
'中间层多语言处理函数
Public Function LoadMKDString(ByVal strWord As String, ByVal strLanguageID As String) As String
    Dim tempErrNumber As Long
    Dim tempErrDescripton As String
    Dim tempErrHelpContext As Variant
    Dim tempErrhelpfile As Variant
    Dim tempErrSource As String
    
    '临时保存传入的Err信息
    If Err.Number <> 0 Then
        tempErrDescripton = Err.Description
        tempErrHelpContext = Err.HelpContext
        tempErrhelpfile = Err.HelpFile
        tempErrSource = Err.Source
        tempErrNumber = Err.Number
    End If
    
    On Error GoTo HANDLEERROR
    Dim LanguageID As String
    
    LanguageID = strLanguageID
    If UCase(LanguageID) <> "CHS" Then
        If (g_objResLoader Is Nothing) Then
            Set g_objResLoader = CreateObject("K3RESLOADER.Loader")
        End If
        '资源文件目录
        'g_objResLoader.ResDirectory = App.Path
        If UCase(g_objResLoader.LanguageID) <> UCase(LanguageID) Then
            g_objResLoader.Unload
            Set g_objResLoader = Nothing
            Set g_objResLoader = CreateObject("K3RESLOADER.Loader")
            g_objResLoader.LanguageID = LanguageID
             '资源文件本名
            g_objResLoader.ResFileBaseName = "K3INDUSTRY"
        End If
        
        '语言编号
        
        strWord = Trim(strWord)
        If InStr(1, strWord, "^|^") > 0 Then
           LoadMKDString = LoadspecialMKDString(strWord, strLanguageID) '存储过程处理，调用LoadspecialMKDString
        Else
           LoadMKDString = g_objResLoader.LoadString(strWord)
        End If
    Else
        strWord = Trim(strWord)
        If InStr(1, strWord, "^|^") > 0 Then
           LoadMKDString = LoadspecialMKDString(strWord, strLanguageID) '存储过程处理，调用LoadspecialMKDString
        Else
           LoadMKDString = strWord
        End If
    End If
    If tempErrNumber <> 0 Then
        Err.Description = tempErrDescripton
        Err.HelpContext = tempErrHelpContext
        Err.HelpFile = tempErrhelpfile
        Err.Source = tempErrSource
        Err.Number = tempErrNumber
    End If
    Exit Function
HANDLEERROR:
        LoadMKDString = strWord
        '恢复传入Err信息
        If tempErrNumber <> 0 Then
            Err.Description = tempErrDescripton
            Err.HelpContext = tempErrHelpContext
            Err.HelpFile = tempErrhelpfile
            Err.Source = tempErrSource
            Err.Number = tempErrNumber
        End If
        'Log it
End Function
'***********************************
'中间层多语言处理增强函数，返回英文时可以控制单词前后添加空格
Public Function LoadMKDStringEx(ByVal strGBText As String, ByVal LanguageID As String, _
                            Optional lEnAppendBlank As EnAppendBlank = EnAppendBlank_NULL) As String
    Dim tempErrNumber As Long
    Dim tempErrDescripton As String
    Dim tempErrHelpContext As Variant
    Dim tempErrhelpfile As Variant
    Dim tempErrSource As String
    
    Dim strWords As String  '裁掉前后空格的字符
    Dim strPreBlank As String   '前面的空格
    Dim strPostBlank As String  '后面的空格
        
    '临时保存传入的Err信息
    If Err.Number <> 0 Then
        tempErrDescripton = Err.Description
        tempErrHelpContext = Err.HelpContext
        tempErrhelpfile = Err.HelpFile
        tempErrSource = Err.Source
        tempErrNumber = Err.Number
    End If
    
    On Error GoTo errhandler
    
    If g_objResLoader Is Nothing Then
        Set g_objResLoader = CreateObject("K3RESLOADER.Loader")
    End If
    
    If UCase(g_objResLoader.ResFileBaseName) <> UCase("K3INDUSTRY") Then
        g_objResLoader.ResFileBaseName = "K3INDUSTRY"
    End If
    
    If g_objResLoader.LanguageID <> LanguageID Then
        g_objResLoader.LanguageID = LanguageID
    End If
    
    '如果有特殊字符，调用LoadspecialKDString
    If InStr(1, strGBText, "^|^") > 0 Then
        LoadMKDStringEx = LoadspecialMKDString(strGBText, LanguageID)
        Exit Function
    Else
        '如果是简体中文就不需要调用资源文件
        If UCase(LanguageID) = UCase("chs") Then
            LoadMKDStringEx = strGBText
        Else
            strWords = Trim(strGBText)
            strPreBlank = VBA.Left$(strGBText, VBA.InStr(1, strGBText, strWords) - 1)
            strPostBlank = VBA.Right$(strGBText, Len(strGBText) - Len(strPreBlank) - Len(strWords))
            
            '需要追加前后空格
            If UCase(LanguageID) = UCase("En") Then
                Select Case lEnAppendBlank
                Case EnAppendBlank_NULL
                Case EnAppendBlank_PREV
                    If Len(strPreBlank) = 0 Then
                        strPreBlank = " "
                    End If
                Case EnAppendBlank_POST
                    If Len(strPostBlank) = 0 Then
                        strPostBlank = " "
                    End If
                Case EnAppendBlank_BOTH
                    If Len(strPreBlank) = 0 Then
                        strPreBlank = " "
                    End If
                    If Len(strPostBlank) = 0 Then
                        strPostBlank = " "
                    End If
                End Select
            End If
            
            LoadMKDStringEx = strPreBlank & g_objResLoader.LoadString(strWords) & strPostBlank
        End If
    End If
    
    If tempErrNumber <> 0 Then
        Err.Description = tempErrDescripton
        Err.HelpContext = tempErrHelpContext
        Err.HelpFile = tempErrhelpfile
        Err.Source = tempErrSource
        Err.Number = tempErrNumber
    End If
    
    Exit Function
    
errhandler:
    'LoadMKDStringEx = "[×]" & strGBText
    LoadMKDStringEx = strGBText
    
    '发版注销
'    Call LogResLoaderErr(strGBText)  '记录翻译错误日志
    '恢复传入Err信息
    If tempErrNumber <> 0 Then
        Err.Description = tempErrDescripton
        Err.HelpContext = tempErrHelpContext
        Err.HelpFile = tempErrhelpfile
        Err.Source = tempErrSource
        Err.Number = tempErrNumber
    End If
End Function


'多语言处理－模板字段替换
Public Function GetLanguageFieldName(ByVal strFieldName As String, ByVal strLanguageID As String) As String
On Error GoTo H_Error
    Select Case UCase(strLanguageID)
        Case "CHT"
            GetLanguageFieldName = strFieldName & "_CHT"
        Case "EN"
            GetLanguageFieldName = strFieldName & "_EN"
        Case Else
            GetLanguageFieldName = strFieldName
    End Select
    Exit Function
H_Error:
    GetLanguageFieldName = strFieldName
End Function


'***********************************************************
'存储过程处理
Public Function LoadspecialMKDString(ByVal strWord As String, ByVal LanguageID As String) As String
     Dim Vargb As Variant
     Dim i As Long
     Dim StrCh As String
     Vargb = Split(strWord, "^|^")
     StrCh = ""
     For i = 0 To UBound(Vargb)
         If Mid(Vargb(i), 1, 3) = "~$~" Then
            Vargb(i) = Mid(Vargb(i), 4)
            StrCh = StrCh & LoadMKDString(Vargb(i), LanguageID)
         Else
            StrCh = StrCh & Vargb(i)
         End If
     Next
     LoadspecialMKDString = StrCh
End Function





