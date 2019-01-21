Attribute VB_Name = "modPub"
Option Explicit

Public g_RetailerItemClassID As Long
Public strMappingFileName As String

Public Function ExecSql(ByVal ssql As String) As Object
    Dim obj As Object
    Dim rs As Object

    Set obj = CreateObject("BillDataAccess.GetData")
    Set rs = obj.ExecuteSQL(MMTS.PropsString, ssql)
    Set obj = Nothing
    Set ExecSql = rs
    Set rs = Nothing
End Function


Public Function UpdateSQL(ByVal ssql As String) As Boolean
    Dim obj As Object
    UpdateSQL = False
    Set obj = CreateObject("BillDataAccess.GetData")
    obj.ExecuteSQL MMTS.PropsString, ssql
    Set obj = Nothing
    UpdateSQL = True
End Function

Public Function ExecuteSQL(sql As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim CNN As Object
    Set CNN = CreateObject("K3Connection.AppConnection")
    Set rs = CNN.Execute(sql)
    Set ExecuteSQL = rs
    Set CNN = Nothing
    Set rs = Nothing
End Function

Public Function CNulls(ByVal v, ByVal default) As Variant
    If IsNull(v) Then
        CNulls = default
    Else
        CNulls = v
    End If
End Function

'控制文本框只能输入字母数字
Public Function Alphabet_Digit_Only(ByVal KeyAscii As Integer) As Integer
    Select Case KeyAscii
        Case 8, 9, 13, &H30 To &H39, Asc("A") To Asc("Z"), Asc("a") To Asc("z")
            Alphabet_Digit_Only = KeyAscii
 
        Case Else
            Alphabet_Digit_Only = 0
    End Select
End Function


'转换连接字符串
Public Function TransfersDsn(ByVal strCatalogName As String, ByVal sDsn As String) As String
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim lStr As String
    Dim rStr As String
    Dim mStr As String
    Dim strDest As String
    
    lStr = Left(sDsn, InStr(1, sDsn, "Catalog") - 1)
    rStr = Right(sDsn, Len(sDsn) - InStr(1, sDsn, "}") + 1)
    mStr = "Catalog=" & strCatalogName
    strDest = lStr & mStr & rStr
    TransfersDsn = strDest
End Function
