Attribute VB_Name = "modPub"
Option Explicit

Public strMappingFileName As String

Public Sub InsertLog(ByVal strLog As String)
    Dim txtFile As Object
    Dim fso As New FileSystemObject
    Dim filePath As String
    
    Dim strDateTimeToString
    
    strDateTimeToString = Format(CStr(Now), "yyyy-mm-dd hh.mm.ss")
        
    If fso.FolderExists(App.Path & "\log") = False Then
        fso.CreateFolder (App.Path & "\log")
    End If
    
    filePath = App.Path & "\log\" & strDateTimeToString & ".txt"
    
    If Dir(filePath) = "" Then '判断文件是否存在，不存在就创建，存在就不创建
        Open filePath For Append As #1
        Close #1
    End If
    
    Set txtFile = fso.OpenTextFile(filePath, 8, True, 0)
    
    txtFile.WriteLine strLog
    
    txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing
    
   'Shell "notepad " & filePath, vbNormalFocus
    
End Sub


Public Function ExecuteSql(ByVal strSQL As String) As adodb.Recordset
    
    Dim MOBJ As Object
    Set MOBJ = CreateObject("BillDataAccess.GetData")
    
    Dim rs As adodb.Recordset
    Set rs = MOBJ.ExecuteSql(MMTS.PropsString, strSQL)
    
    Set ExecuteSql = rs
    Set rs = Nothing
    Set MOBJ = Nothing
End Function

Public Function GetAnyRecordSet(ByVal ConStr As String, ByVal strSQL As String) As adodb.Recordset
    
    Dim MOBJ As Object
    Dim rs As adodb.Recordset
    
    Set MOBJ = CreateObject("K3MAppConnection.AppConnection")
    Set rs = MOBJ.GetAnyRecordSet(ConStr, strSQL)
    
    Set GetAnyRecordSet = rs
    Set rs = Nothing
    Set MOBJ = Nothing
    
End Function


Public Sub ExecuteSQLWithTrans(ByVal strSQL As String)
    Dim MOBJ As Object
    
    Set MOBJ = CreateObject("K3MAppConnection.AppConnection")
    MParse.ParseString MMTS.PropsString
    MOBJ.Execute MParse.ConStr, strSQL
    Set MOBJ = Nothing

End Sub


Public Function ExeSqlWithCnn(ByVal strCnn As String, ByVal strSQL As String) As adodb.Recordset
'    Dim cnn As New ADODB.Connection
    Dim rs As New adodb.Recordset
On Error GoTo hrr
'    cnn.Open strCnn
    rs.CursorLocation = adUseClient
    rs.Open strSQL, strCnn, adOpenStatic, adLockOptimistic
'    Set rs = cnn.Execute(strSql)
    Set ExeSqlWithCnn = rs
    Set rs = Nothing
'    cnn.Close
'    Set cnn = Nothing
    Exit Function
hrr:
    Set rs = Nothing
'    cnn.Close
'    Set cnn = Nothing
'    MsgBox Err.Description
End Function



Public Function ChangeStr(ByVal str As String) As String
    Dim i As Integer
    
    If Len(str) < 4000 Then
        ChangeStr = str
        Exit Function
    End If
    
    Do While Len(str) >= 4000
        i = InStrRev(str, vbCrLf)
        str = CStr(Left(str, i - 1))
    Loop
    
    ChangeStr = str
    
End Function

















