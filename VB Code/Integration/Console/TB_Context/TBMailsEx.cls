VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBMailsEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ThisSize As Long
Private ThisKeys() As String
Private ThisTo() As TBMails
Private ThisCC() As TBMails

Private Sub Class_Terminate()
    Erase ThisTo
    Erase ThisCC
End Sub

Public Property Get Key(ByVal Index As Long) As String
    Key = ThisKeys(Index)
End Property

Public Property Get ToMail(ByVal Index As Long) As TBMails
    Set ToMail = ThisTo(Index)
End Property
Public Property Set ToMail(ByVal Index As Long, mails As TBMails)
    Set ThisTo(Index) = mails
End Property

Public Property Get CCMail(ByVal Index As Long) As TBMails
    Set CCMail = ThisCC(Index)
End Property
Public Property Set CCMail(ByVal Index As Long, mails As TBMails)
    Set ThisCC(Index) = mails
End Property

Public Property Get Size() As Long
    Size = ThisSize
End Property

Public Function Lookup(sKey As String) As Long
    Dim i As Long
    For i = 1 To ThisSize
        If ThisKeys(i - 1) = sKey Then
            Lookup = i - 1
            Exit Function
        End If
    Next i
    
    Lookup = mParam.CONST_ERROR_INDEX
End Function

Public Function AddToMail(ByVal sKey As String, mails As TBMails)
    Dim i As Long
    i = Lookup(sKey)
    If i > -1 Then
        Set ThisTo(i) = mails
    Else
        ReDim Preserve ThisKeys(ThisSize)
        ReDim Preserve ThisTo(ThisSize)
        ReDim Preserve ThisCC(ThisSize)
        ThisKeys(ThisSize) = sKey
        Set ThisTo(ThisSize) = mails
        ThisSize = ThisSize + 1
    End If
    
    AddToMail = True
End Function

Public Function AddCCMail(ByVal sKey As String, mails As TBMails)
    Dim i As Long
    i = Lookup(sKey)
    If i > -1 Then
        Set ThisCC(i) = mails
    Else
        ReDim Preserve ThisKeys(ThisSize)
        ReDim Preserve ThisTo(ThisSize)
        ReDim Preserve ThisCC(ThisSize)
        ThisKeys(ThisSize) = sKey
        Set ThisCC(ThisSize) = mails
        ThisSize = ThisSize + 1
    End If
    
    AddCCMail = True
End Function

