VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLisence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------
'AUTHOR:        Zhang JianJian
'DATE:          2013-11-05
'LAST UPDATE:   2013-11-05
'DESCRIPTION:   加解密算法(简单文本)，对称密钥，不支持数据流
'--------------------------------------------------

Private Const ERROR_NUMBER_INVALIDKEY As Long = 30001
Private Const ERROR_TEXT_INVALIDKEY As String = "无效的密钥"

Private KEY_TEXT As String
Private KEY_LEN As Long
Private CODE_ROUND As ENUM_CODEROUND
Private CODE_LEVEL As ENUM_CODELEVEL
Private CODE_FASHION As ENUM_CODEFASHION

Public Enum ENUM_CODEROUND
    L1 = 16
    L2 = 32
    L3 = 64
End Enum

Public Enum ENUM_CODELEVEL
    L1 = 16
    L2 = 32
    L3 = 64
    L4 = 128
    L5 = 256
End Enum

Public Enum ENUM_CODEFASHION
    ASCII
    UNICODE
    UTF8
End Enum

'明钥
Public Property Get KeyText() As String
    KeyText = KEY_TEXT
End Property
Public Property Let KeyText(ByVal text As String)
    KEY_TEXT = text
    KEY_LEN = Len(KEY_TEXT)
End Property

'密钥长度
Public Property Get KeyLength() As Long
    KeyLength = KEY_LEN
End Property

'轮询次数
Public Property Get CodeRound() As ENUM_CODEROUND
    CodeRound = CODE_ROUND
End Property
Public Property Let CodeRound(ByVal round As ENUM_CODEROUND)
    CODE_ROUND = round
End Property

'加密深度
Public Property Get CodeLevel() As ENUM_CODELEVEL
    CodeLevel = CODE_LEVEL
End Property
Public Property Let CodeLevel(ByVal level As ENUM_CODELEVEL)
    CODE_LEVEL = level
End Property

'国际化/本地化
Public Property Get CodeFashion() As ENUM_CODEFASHION
    CodeFashion = CODE_FASHION
End Property
Public Property Let CodeFashion(ByVal fashion As ENUM_CODEFASHION)
    CODE_FASHION = fashion
End Property

Private Sub Class_Initialize()
    KEY_TEXT = "FUN4.ORG All Rights Reserved!"
    CODE_ROUND = ENUM_CODEROUND.L1
    CODE_LEVEL = ENUM_CODELEVEL.L3
    CODE_FASHION = ENUM_CODEFASHION.UNICODE
End Sub

'简单文本加密
Public Function Encode(ByVal text As String) As String
    Dim vKey() As Byte
    Dim vText() As Byte
    
    '数据片段
    Dim lPage As Long
    Dim vPage(KEY_LEN) As Byte
    
    '当前轮询次数
    Dim lRound As Long
    
    vKey = InitKeys(KEY_TEXT)
    vText = InitTexts(text)
    
    While lRound < CODE_ROUND
        vKey = GetKeys(vKey, lRound)
        
    Wend
End Function

'简单文本解密
Public Function Decode(ByVal text As String) As String
    
End Function

'初始化密钥，符合加密深度
Private Function InitKeys(ByVal sKey As String) As Byte()
    Dim i As Long, j As Long, z As Long
    Dim vKey() As Byte
    Dim vNewKey(CODE_LEVEL - 1) As Byte
    
    If LenB(sKey) = 0 Then
        Err.Raise ERROR_NUMBER_INVALIDKEY, "InitKeys", ERROR_TEXT_INVALIDKEY
    End If
    
    vKey = Serialize(sKey)
    j = UBound(vKey)
    
    While i < CODE_LEVEL
        vNewKey(i) = vKey(z)
        
        If j = z Then
            z = 0
        Else
            z = z + 1
        End If
    Wend

    InitKeys = vNewKey
    
    Erase vKey
    Erase vNewKey
End Function

'初始化明文
Private Function InitTexts(ByVal sText As String) As Byte()
    Dim vText() As Byte
    
    vText = Serialize(vText)
    If UBound(vText) < CODE_LEVEL Then
        ReDim Preserve vText(CODE_LEVEL)
    End If
    InitTexts = vText
    Erase vText
End Function

'密钥生成器
Private Function GetKeys(ByVal vKey As Variant, ByVal lRound As Long) As Byte()
    
End Function

Private Function GetPages(vText As Variant, ByVal lPage As Long) As Byte()
    
End Function

'国际化，暂不处理
Public Function Serialize(ByVal text As String) As Byte()
    Dim v() As Byte
    v = text
    Serialize = v
    Erase v
End Function

'本地化，暂不处理
Public Function DeSerialize(ByVal vText As Variant) As String
    Dim s As String
    s = vText
    DeSerialize = s
    Erase vText
End Function

