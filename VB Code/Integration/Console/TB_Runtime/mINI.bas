Attribute VB_Name = "mINI"
Option Explicit

Public Config_INI_Patch As String

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub SetIniValue(ByVal Section As String, ByVal Key As String, ByVal Value As String)
    WritePrivateProfileString Section, Key, Value, Config_INI_Patch
End Sub

Public Function GetIniValue(ByVal Section As String, ByVal Key As String, Optional ByVal DefValue As String = "") As String
    Dim sTemp As String
    Dim lRet As String
    Dim i As Integer
    Dim sValue As String

    sTemp = String$(1024, Chr(32))
    lRet = GetPrivateProfileString(Section, Key, "", sTemp, Len(sTemp), Config_INI_Patch)

    sTemp = Trim(sTemp)
    
    For i = 1 To Len(sTemp)
        If Asc(Mid(sTemp, i, 1)) <> 0 Then
            sValue = sValue + Mid(sTemp, i, 1)
        End If
    Next i
    
    GetIniValue = IIf(sValue = "", DefValue, sValue)
End Function

Public Function DelIniKey(ByVal Section As String, ByVal Key As String) As Boolean
    WritePrivateProfileString Section, Key, 0&, Config_INI_Patch
End Function

Public Function DelIniSec(ByVal Section As String) As Boolean
    WritePrivateProfileString Section, 0&, "", Config_INI_Patch
End Function

