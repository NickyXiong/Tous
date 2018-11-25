VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lisence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Encode(ByVal s As String, ByVal k As String) As String
    Dim vKey() As Byte
    Dim vText() As Byte
    
    vKey = k
    vText = s
    
    Convert vText
    TextXOR vKey, vText
    
    Encode = vText
End Function

Public Function Decode(ByVal s As String, ByVal k As String) As String
    Dim vKey() As Byte
    Dim vText() As Byte
    
    vKey = k
    vText = s
    
    TextXOR vKey, vText
    Convert vText

    Decode = vText
End Function

Private Sub Convert(vText As Variant)
    Dim i As Long, j As Long, l As Long
    Dim b As Byte
    
    j = UBound(vText)
    l = j / 2 - 1
    
    For i = 0 To l
        b = vText(i)
        vText(i) = vText(j - i)
        vText(j - i) = b
    Next i
End Sub

Private Sub TextXOR(ByVal vKey As Variant, vText As Variant)
    Dim i As Long, j As Long, l As Long, z As Long
    
    j = UBound(vText)
    z = UBound(vKey)
    
    For i = 0 To j
        vText(i) = vText(i) Xor vKey(i Mod z)
    Next i
End Sub
