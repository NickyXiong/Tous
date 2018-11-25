Attribute VB_Name = "modPub"



Public Function ExecSQL(ByVal sdsn As String, ByVal strSql As String) As ADODB.Recordset
    Dim cn As ADODB.Connection
    InitDataEnv sdsn
    Set cn = datasource.Connection
    Set ExecSQL = cn.Execute(strSql)
End Function
Private Function InitDataEnv(ByVal sToParse As String) As Boolean
    Dim m_oParse As Object
    Set m_oParse = New CParse
    If m_oParse.ParseString(sToParse) Then
        Set datasource = New CDataSource
        Set datasource.ParseObject = m_oParse
    Else
        Err.Raise EBS_E_TypeMismatch, "ParseString"
    End If
End Function
