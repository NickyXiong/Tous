VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Kingdee Enterprise Business Objects
' Copyright (C) 1995-1998 Kingdee Corporation
' All rights reserved

Option Explicit
Public ParseObject As CParse
Public Property Get Connection() As ADODB.Connection
    Dim mConnection As ADODB.Connection
    Set mConnection = New ADODB.Connection
    mConnection.Open ConnectString
    mConnection.Execute "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
    mConnection.CursorLocation = adUseClient
    mConnection.CommandTimeout = 0
    Set Connection = mConnection
End Property

Public Property Get ConnectionSrv() As ADODB.Connection
    Dim mConnection As ADODB.Connection
    Set mConnection = New ADODB.Connection
    mConnection.Open ConnectString
    mConnection.Execute "SET TRANSACTION ISOLATION LEVEL READ COMMITTED"
    mConnection.Execute "SET NOCOUNT ON"
    mConnection.Execute "DBCC TRACEOFF"
    mConnection.CursorLocation = adUseServer
    mConnection.CommandTimeout = 0
    Set ConnectionSrv = mConnection
    Set mConnection = Nothing
End Property

Private Function OpenConnection(Optional bRaiseError = False) As Boolean
    OpenConnection = False
    On Error GoTo ErrorHandler
    
    ' retrieve connect string from system registry
    Dim strConn As String
''''    strConn = KFO.Reg.GetSystemSetting("EBO", "Configuration", "DataSource")
''''    If strConn = "" Then
''''        Err.Raise &H80004005, "EBSBase.DataSource.OpenConnection" ' E_FAIL
''''        Exit Function
''''    End If
    strConn = ParseObject.ConStr
    
    ' Test connection usability
    Dim cn As New ADODB.Connection
    cn.Open strConn
    
    ' The connection has successfully opened, save connection
    ' properties in shared property group
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProp As SharedProperty
    Dim bExists As Boolean
    
    Set spmMgr = CreateObject("MtxSpm.SharedPropertyGroupManager.1")
''''    Set spmGroup = spmMgr.CreatePropertyGroup("DataSource", LockSetGet, Process, bExists)
    Dim strGroup As String
    strGroup = ParseObject.ConDBMSName & ParseObject.ConDBMSVersion & "DataSource"
    Set spmGroup = spmMgr.CreatePropertyGroup(strGroup, LockSetGet, Process, bExists)
    
    ' Not all properties are supported by data source, so ignore
    ' errors when failed to retrieve connection properties
    Dim var As Variant
    
    On Error Resume Next
    cn.Properties.Refresh
    
    Set spmProp = spmGroup.CreateProperty("DBMSName", bExists)
    var = cn.Properties("DBMS Name").Value
    If IsNull(var) Or IsEmpty(var) Then var = vbNullString
    spmProp.Value = var
    
    Set spmProp = spmGroup.CreateProperty("DBMSVersion", bExists)
    var = cn.Properties("DBMS Version").Value
    If IsNull(var) Or IsEmpty(var) Then var = vbNullString
    spmProp.Value = var
    
    Set spmProp = spmGroup.CreateProperty("MultipleResults", bExists)
    var = cn.Properties("Multiple Results").Value
    If IsNull(var) Or IsEmpty(var) Then var = False
    spmProp.Value = var
    
    Set spmProp = spmGroup.CreateProperty("LikeEscapeClause", bExists)
    var = cn.Properties("Like Escape Clause").Value
    If IsNull(var) Or IsEmpty(var) Then var = "N"
    spmProp.Value = (var = "Y")
    
    Set spmProp = spmGroup.CreateProperty("IEF", bExists)
    var = cn.Properties("Integrity Enhancement Facility").Value
    If IsNull(var) Or IsEmpty(var) Then var = False
    spmProp.Value = var
    
    Set spmProp = spmGroup.CreateProperty("NumericFunctions", bExists)
    spmProp.Value = cn.Properties("Numeric Functions").Value
    Set spmProp = spmGroup.CreateProperty("StringFunctions", bExists)
    spmProp.Value = cn.Properties("String Functions").Value
    Set spmProp = spmGroup.CreateProperty("SystemFunctions", bExists)
    spmProp.Value = cn.Properties("System Functions").Value
    Set spmProp = spmGroup.CreateProperty("DateTimeFunctions", bExists)
    spmProp.Value = cn.Properties("Time/Date Functions").Value
    
    On Error GoTo ErrorHandler
    
    ' Retrieve all supported procedures
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT FProcName FROM t_SupportedProc", cn, _
        adOpenForwardOnly, adLockReadOnly, adCmdText

    Dim dict As New KFO.Dictionary
    Do Until rs.EOF
        dict(rs!FProcName) = True
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Set spmProp = spmGroup.CreateProperty("SupportedProc", bExists)
    spmProp.Value = dict
    
    ' Cache connect string and set connected flag
    Set spmProp = spmGroup.CreateProperty("ConnectString", bExists)
    spmProp.Value = strConn
    Set spmProp = spmGroup.CreateProperty("Connected", bExists)
    spmProp.Value = True
    
    cn.Close
    OpenConnection = True
    Exit Function
    
ErrorHandler:
    OpenConnection = False
    If bRaiseError Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Private Function GetConnectionProperty(strName As String, _
    Optional bRaiseError As Boolean = False) As Variant
    
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProp As SharedProperty
    Dim bExists As Boolean
    
    Set spmMgr = CreateObject("MTxSpm.SharedPropertyGroupManager.1")
    '------Set spmGroup = spmMgr.CreatePropertyGroup("DataSource", LockSetGet, Process, bExists)
    Dim strGroup As String
    strGroup = ParseObject.ConDBMSName & ParseObject.ConDBMSVersion & "DataSource"
    Set spmGroup = spmMgr.CreatePropertyGroup(strGroup, LockSetGet, Process, bExists)
    
    Set spmProp = spmGroup.CreateProperty("Connected", bExists)
    If spmProp.Value <> True Then
        If Not OpenConnection(bRaiseError) Then
            Exit Function
        End If
    End If
    
    Set spmProp = spmGroup.Property(strName)
    If IsObject(spmProp.Value) Then
        Set GetConnectionProperty = spmProp.Value
    Else
        GetConnectionProperty = spmProp.Value
    End If
End Function

Public Property Get ConnectString() As String
    '---ConnectString = GetConnectionProperty("ConnectString", True)
    ConnectString = ParseObject.ConStr
End Property

Public Property Get DBMSName() As String
    '----DBMSName = GetConnectionProperty("DBMSName")
    DBMSName = ParseObject.ConDBMSName
End Property

Public Property Get DBMSVersion() As String
    DBMSVersion = GetConnectionProperty("DBMSVersion")
End Property

Public Property Get MultipleResults() As Boolean
    MultipleResults = GetConnectionProperty("MultipleResults")
End Property

Public Property Get LikeEscapeClause() As Boolean
    LikeEscapeClause = GetConnectionProperty("LikeEscapeClause")
End Property

Public Property Get IEF() As Boolean
    IEF = GetConnectionProperty("IEF")
End Property

Public Property Get NumericFunctions() As Long
    NumericFunctions = GetConnectionProperty("NumericFunctions")
End Property

Public Property Get StringFunctions() As Long
    StringFunctions = GetConnectionProperty("StringFunctions")
End Property

Public Property Get SystemFunctions() As Long
    SystemFunctions = GetConnectionProperty("SystemFunctions")
End Property

Public Property Get DateTimeFunctions() As Long
    DateTimeFunctions = GetConnectionProperty("DateTimeFunctions")
End Property

Public Function ProcSupported(ProcName As String) As Boolean
    Dim dict As KFO.Dictionary
    Set dict = GetConnectionProperty("SupportedProc", True)
    ProcSupported = dict.Lookup(ProcName)
End Function

Public Sub LoadDataTypeInfo(spmGroup As SharedPropertyGroup)
    Dim spmProp As SharedProperty
    Dim bExists As Boolean
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "select * from t_DataTypeInfo", ConnectString
    Do Until rs.EOF
        Set spmProp = spmGroup.CreatePropertyByPosition(rs("FDataType"), bExists)
        spmProp.Value = PackageRecord(rs)
        rs.MoveNext
    Loop
End Sub

Private Function GetDataTypeInfo(ByVal DataType As ADODB.DataTypeEnum) As KFO.Dictionary
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProp As SharedProperty
    Dim bExists As Boolean
    
    Set spmMgr = CreateObject("MTxSpm.SharedPropertyGroupManager.1")
    Dim strGroup As String
    strGroup = ParseObject.ConDBMSName & ParseObject.ConDBMSVersion & "DataTypeInfo"
    Set spmGroup = spmMgr.CreatePropertyGroup(strGroup, LockSetGet, Process, bExists)
    
    If Not bExists Then
        LoadDataTypeInfo spmGroup
    End If
    
    Set spmProp = spmGroup.PropertyByPosition(DataType)
    Set GetDataTypeInfo = spmProp.Value
End Function

Public Function GetTypeName( _
    ByVal DataType As ADODB.DataTypeEnum, _
    Optional ByRef Precision As Long = 0, _
    Optional ByRef nScale As Integer = 0, _
    Optional ByRef ActualType As Integer, _
    Optional ByRef ActualSize As Long, _
    Optional ByRef BindType As Integer) As String

    Dim tinfo As KFO.Dictionary
    Dim MaxPrecision As Variant
    Dim MinScale As Variant
    Dim MaxScale As Variant
    Dim TypeName As String
    
    Dim strTypeName As String
    
    Set tinfo = GetDataTypeInfo(DataType)
    MaxPrecision = tinfo("FMaxPrecision")
    MinScale = tinfo("FMinScale")
    MaxScale = tinfo("FMaxScale")
    TypeName = tinfo("FTypeName")
    
    If Not IsNull(MaxPrecision) Then
        If MaxPrecision < Precision Then
            ' The type precision is not big enough as
            ' we requested
            Exit Function
        End If
    End If
        
    If Not IsNull(MaxScale) Then
        If nScale > MaxScale Then
            nScale = MaxScale
        End If
    End If
    If Not IsNull(MinScale) Then
        If nScale < MinScale Then
            nScale = MinScale
        End If
    End If
        
    ' If a precision is needed?
    If Not IsNull(MaxPrecision) Then
        ' Usually most type name that requires a precision
        ' and scale will be returned by data source as just
        ' the names without the parens. For example: NUMERIC.
        ' Sometimes, though, the parens may be returned
        ' within the type name (CHAR() FOR BIT DATA). We
        ' really need to search for parens within type name
        ' and place the parameters within them. If there are
        ' no parens within the type name, we can just add
        ' them at the end.

        Dim nParenStart As Integer, nParenEnd As Integer
        
        nParenStart = InStr(TypeName, "(")
        If nParenStart > 0 Then
            nParenEnd = InStr(TypeName, ")")
            If nParenEnd = 0 Then
                nParenEnd = nParenStart
            End If
            nParenEnd = nParenEnd + 1
                
            ' Copy the name before and include open parens
            strTypeName = Left(TypeName, nParenStart)
                
            ' Copy the precision to
            strTypeName = strTypeName & Precision
                
            ' If a scale is needed?
            If Not IsNull(MaxScale) Then
                strTypeName = strTypeName & "," & nScale
            End If
                
            strTypeName = strTypeName + ")"
                
            ' Copy the type name after close parens
            strTypeName = strTypeName + Mid(TypeName, nParenEnd)
        Else
            ' No parens found in type name, add precision
            ' and scale at the end of type name.
            strTypeName = TypeName + "(" & Precision
            If Not IsNull(MaxScale) Then
                strTypeName = strTypeName + "," & nScale
            End If
            strTypeName = strTypeName + ")"
        End If
    Else
        ' Just copy the type name
        strTypeName = TypeName
            
        ' Neither precision nor scale is needed.
        Precision = 0
        nScale = 0
    End If
        
    ActualType = tinfo("FActualType")
    BindType = tinfo("FBindType")
    If Not IsNull(tinfo("FActualSize")) Then
        ActualSize = tinfo("FActualSize")
    Else
        ActualSize = Precision
    End If

    GetTypeName = strTypeName
End Function

Public Function GetActualType( _
    ByVal DataType As ADODB.DataTypeEnum, _
    Optional ByRef Precision As Long = 0, _
    Optional ByRef nScale As Integer = 0, _
    Optional ByRef ActualSize As Long, _
    Optional ByRef BindType As Integer) As Integer
    
    Dim tinfo As KFO.Dictionary
    Set tinfo = GetDataTypeInfo(DataType)
    
    If Not IsNull(tinfo("FMaxPrecision")) Then
        If tinfo("FMaxPrecision") < Precision Then
            ' The type precision is not big enough as
            ' we requested
            Exit Function
        End If
    End If
        
    If Not IsNull(tinfo("FMaxScale")) Then
        If nScale > tinfo("FMaxScale") Then
            nScale = tinfo("FMaxScale")
        End If
    End If
    If Not IsNull(tinfo("FMinScale")) Then
        If nScale < tinfo("FMinScale") Then
            nScale = tinfo("FMinScale")
        End If
    End If
    
    GetActualType = tinfo("FActualType")
    BindType = tinfo("FBindType")
    If Not IsNull(tinfo("FActualSize")) Then
        ActualSize = tinfo("FActualSize")
    Else
        ActualSize = Precision
    End If
End Function

Public Function GetBindType(ByVal DataType As ADODB.DataTypeEnum) As Integer
    Dim tinfo As KFO.Dictionary
    Set tinfo = GetDataTypeInfo(DataType)
    GetBindType = tinfo("FBindType")
End Function

Public Property Get IsOracle() As Boolean
    IsOracle = DBMSName() = "Oracle"
End Property

'Add by flin 2001.11.29
'-----------------------------------------------------------------------------------------------
'描述:处理批量执行SQL时的ADO错误
'
'入口参数:cn = 需要检查的数据库连接对象
'         bRaiseError = 是否触发VB Runtime Error，可选，默认为true
'出口参数:无
'
'返回值:为True表示存在错误
'
'算法说明:无
'-----------------------------------------------------------------------------
Public Function HasADOError(cn As ADODB.Connection, Optional bRaiseError As Boolean = True) As Boolean
Dim oADOError As ADODB.Error

'KDErrorHandle.CheckADOError对象，该对象的IsError(oADOError)函数用于检测某指定的oADOError是否是错误
'某些情况下Provider的警告信息也会放入Errors集合中，但不影响数据的正确性与一致性
'这些特殊的情况判断可以放入IsError函数的处理中
Dim oCheckError As Object

'Errors集合中第一个错误号，用于VB Run Time的错误号
Dim lFirstErrorNumber As Long
'Errors所有错误的集合
Dim sErrorInfo As String
'是否有错误存在的指示变量
Dim bHasError As Boolean
If cn.Errors.Count > 0 Then
    bHasError = False
    
    '用于KDErrorHandle.CheckADOError无法创建时的处理
    'KDErrorHandle.CheckADOError无法创建时处理所有ADO错误
    On Error Resume Next
    Set oCheckError = CreateObject("KDErrorHandle.CheckADOError")
    If Err.Number <> 0 Then
        lFirstErrorNumber = Err.Number
        sErrorInfo = "Number:" & Err.Number & vbCrLf _
                & "Description:" & Err.Description & vbCrLf
        bHasError = True
    End If
    On Error GoTo 0
    
    For Each oADOError In cn.Errors
        If Not bHasError Then
            If oCheckError.IsError(oADOError) Then
                lFirstErrorNumber = oADOError.Number
                bHasError = True
            End If
        End If
        sErrorInfo = sErrorInfo & "Number:" & oADOError.Number & vbCrLf _
                & "Source:" & oADOError.Source & vbCrLf _
                & "Description:" & oADOError.Description & vbCrLf _
                & "NativeError:" & oADOError.NativeError & vbCrLf _
                & "SQLState:" & oADOError.SQLState & vbCrLf
    Next
    
    If bRaiseError And bHasError Then
        Err.Raise lFirstErrorNumber, , sErrorInfo
    End If
    HasADOError = bHasError
End If
End Function



