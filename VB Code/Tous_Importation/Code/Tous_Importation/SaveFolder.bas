Attribute VB_Name = "SaveFolder"
Declare Function SHGetPathFromIDList Lib "shell32.dll" _
Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
ByVal pszPath As String) As Long

Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type

Function GetDirectory(Optional Msg) As String
Dim bInfo As BROWSEINFO
Dim path As String
Dim r As Long, x As Long, pos As Integer
' Root folder = Desktop
    bInfo.pidlRoot = 0&
' Title in the dialog
If IsMissing(Msg) Then
    bInfo.lpszTitle = "Select a folder."
Else
    bInfo.lpszTitle = Msg
End If
' Type of directory to return
    bInfo.ulFlags = &H1
' Display the dialog
    x = SHBrowseForFolder(bInfo)
' Parse the result
    path = Space$(512)
    r = SHGetPathFromIDList(ByVal x, ByVal path)
If r Then
    pos = InStr(path, Chr$(0))
    GetDirectory = Left(path, pos - 1)
Else
    GetDirectory = ""
End If
End Function


