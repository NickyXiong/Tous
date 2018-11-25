VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim filenumber As Integer
Dim fullfilename As String

Public Sub InitLogWithFileName(ByVal filepath As String, ByVal strfilename As String)
    If filenumber <> 0 Then
        Close #filenumber
    End If
    filenumber = FreeFile
    
    Dim fso As New FileSystemObject

    Dim strDateTimeToString
    
    strDateTimeToString = Format(CStr(Now), "yyyymmddhhmmss")
        
    If fso.FolderExists(filepath) = False Then
        fso.CreateFolder (filepath)
    End If
    
    fullfilename = filepath & "\" & strfilename & " - error log - " & strDateTimeToString & ".txt"
    
   ' If Dir(filepath) = "" Then '判断文件是否存在，不存在就创建，存在就不创建
        Open fullfilename For Output As #filenumber
  '  End If
    
End Sub

'写入日志
'strFileName是被导入的文件的文件名,这里作为输入参数用来构造成日志文件名的一部分
Public Sub WriteLine(ByVal strLog As String)
'    Dim txtFile As Object
'    Dim fso As New FileSystemObject
'    Dim filepath As String
'
'    Dim strDateTimeToString
'
'    strDateTimeToString = Format(CStr(Now), "yyyymmddhhmmss")
'
'    If fso.FolderExists(txtFailure.Text) = False Then
'        fso.CreateFolder (txtFailure.Text)
'    End If
'
'    filepath = txtFailure.Text & "\" & strfilename & " - error log - " & strDateTimeToString & ".txt"
'
'    If Dir(filepath) = "" Then '判断文件是否存在，不存在就创建，存在就不创建
'        Open filepath For Append As #1
'        Close #1
'    End If
'
'    Set txtFile = fso.OpenTextFile(filepath, 8, True, 0)
'
'    txtFile.WriteLine strLog
'
'    txtFile.Close
'    Set txtFile = Nothing
'    Set fso = Nothing
'
   'Shell "notepad " & filePath, vbNormalFocus
    Print #filenumber, strLog
End Sub

Private Sub Class_Terminate()
    If filenumber <> 0 Then
        Close #filenumber
    End If
End Sub

Public Sub CloseFile()
    If filenumber <> 0 Then
        Close #filenumber
        filenumber = 0
    End If
End Sub

Public Sub DeleteFile()
    If Dir(fullfilename) <> "" Then
        Kill fullfilename
    End If
    
End Sub
