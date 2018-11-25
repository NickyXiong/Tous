Attribute VB_Name = "modParse"
Option Explicit


Public Function GetProperty(ByRef coll As Collection, ByVal sPropName As String) As String
    Dim sValue As String
    On Error Resume Next
    sValue = ""
    sValue = coll(sPropName)
    GetProperty = sValue
    Err.Clear
End Function
Public Function ParseString(ByRef coll As Collection, ByVal sToParse As String, Optional ByVal sCompare As String = "=", Optional ByVal sSep As String = ";") As Boolean
Dim sName As String
Dim sValue As String
Set coll = New Collection
    Do
        sName = GetName(sToParse)
        sValue = GetValue(sToParse, sSep)
        If sName <> "" Then
            coll.Add sValue, sName
        Else
            Exit Do
        End If
    Loop
    ParseString = True
End Function
Private Function SearchString(sBeSearch As String, ByVal sFind As String) As String
On Error GoTo Err_SearchString
Dim v As Variant
v = Split(sBeSearch, sFind, 2, vbTextCompare)
Dim lb As Integer, ub As Integer
    lb = LBound(v)
    ub = UBound(v)
    If ub > lb Then
        sBeSearch = v(ub)
        SearchString = v(lb)
    ElseIf ub = lb Then
        sBeSearch = ""
        SearchString = v(ub)
    Else
        sBeSearch = ""
        SearchString = ""
    End If
    Exit Function
Err_SearchString:
    sBeSearch = ""
    SearchString = ""
End Function
Private Function GetName(sBeSearch As String, Optional ByVal sCompare As String = "=") As String
    GetName = SearchString(sBeSearch, sCompare)
    GetName = Trim$(GetName)
End Function
Private Function GetValue(sBeSearch As String, Optional ByVal sSep As String = ";") As String
sBeSearch = Trim$(sBeSearch)
If VBA.Left$(sBeSearch, 1) = "{" Then
    sBeSearch = Mid$(sBeSearch, 2)
    GetValue = SearchString(sBeSearch, "}")
    SearchString sBeSearch, sSep
Else
    GetValue = SearchString(sBeSearch, sSep)
End If
    GetValue = Trim$(GetValue)
End Function
