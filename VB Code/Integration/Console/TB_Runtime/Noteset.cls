VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Noteset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ThisIsGroup As Boolean
Private ThisIndexs As Variant
Private ThisCount As Long
Private ThisData() As Variant

Private ThisPoint As Long
Private ThisIndexCount As Long
Private ThisIndexKey() As String
Private ThisIndexValue() As String

Public Property Get IndexCount() As Long
    If ThisIsGroup Then
        IndexCount = ThisIndexCount
    Else
        IndexCount = ThisCount
    End If
End Property

Public Property Get BOF() As Boolean
    If ThisPoint = -1 Then
        BOF = True
    End If
End Property

Public Property Get EOF() As Boolean
    If ThisPoint >= IndexCount Then
        EOF = True
    End If
End Property

Public Sub LoadData(ByVal locDir As String, Optional ByVal Indexs As Variant, _
                                        Optional ByVal Charset As String = "UTF-8", _
                                        Optional ByVal SplitChar As String = vbTab, _
                                        Optional ByVal LineChar As String = vbCrLf)
    Dim stream As String
    Dim vStream() As String
    Dim vCol() As String
    Dim l As Long, lCount As Long
    Dim lType As Long
    
    Dispose
    
    If IsArray(Indexs) Then
        If UBound(Indexs) > -1 Then
            ThisIndexs = Indexs
            ThisIsGroup = True
            lType = 2
        End If
    Else
        If IsNumeric(Indexs) Then
            If Indexs > 0 Then
                ThisIndexs = Indexs
                ThisIsGroup = True
                lType = 1
            End If
        End If
    End If
    
    stream = LoadAs(locDir, Charset)
    
    If Len(stream) = 0 Then
        ThisCount = 0
        Exit Sub
    End If
    
    vStream = Split(stream, LineChar)
    lCount = UBound(vStream)
    ReDim ThisData(lCount)
    
    If ThisIsGroup Then
        ReDim ThisIndexKey(lCount)
        ReDim ThisIndexValue(lCount)
    End If
    
    For l = 0 To lCount
        If Len(vStream(l)) > 0 Then
            
            vCol = Split(vStream(l), SplitChar)
            
            If ThisIsGroup Then
                If lType = 2 Then
                    AppendByIndexs vCol, l
                Else
                    AppendByIndex vCol(ThisIndexs - 1), l
                End If
            End If
            
            ThisData(ThisCount) = vCol
            ThisCount = ThisCount + 1
            
            Erase vCol
        End If
    Next l
    
    Erase vStream
End Sub

Private Sub AppendByIndexs(vCol As Variant, Row As Long)
    Dim l As Long
    Dim bRet As Boolean
    Dim Key2 As String
    
    For l = 0 To UBound(ThisIndexs)
        Key2 = Key2 & "^!^" & vCol(ThisIndexs(l) - 1)
    Next l
    
    For l = 0 To ThisIndexCount - 1
        If ThisIndexKey(l) = Key2 Then
            ThisIndexValue(l) = ThisIndexValue(l) & "." & Row
            bRet = True
            Exit For
        End If
    Next l
    
    If Not bRet Then
        ThisIndexKey(ThisIndexCount) = Key2
        ThisIndexValue(ThisIndexCount) = Row
        ThisIndexCount = ThisIndexCount + 1
    End If
End Sub

Private Sub AppendByIndex(Key As String, Row As Long)
    Dim l As Long
    Dim bRet As Boolean
    
    For l = 0 To ThisIndexCount - 1
        If ThisIndexKey(l) = Key Then
            ThisIndexValue(l) = ThisIndexValue(l) & "." & Row
            bRet = True
            Exit For
        End If
    Next l
    
    If Not bRet Then
        ThisIndexKey(ThisIndexCount) = Key
        ThisIndexValue(ThisIndexCount) = Row
        ThisIndexCount = ThisIndexCount + 1
    End If
End Sub

Public Function SaveData(ByVal locDir As String, vData As Variant, _
                                        Optional ByVal Charset As String = "GB2312", _
                                        Optional ByVal SplitChar As String = vbTab, _
                                        Optional ByVal LineChar As String = vbCrLf) As Boolean
    Dim stream As String
    Dim i As Long
    Dim j As Long
    
    Dim vRow As Variant
    Dim oStream As ADODB.stream
    
    Set oStream = New ADODB.stream
    
    If Len(Dir(locDir, vbArchive)) > 0 Then
        FileSystem.Kill locDir
    End If
    
    With oStream
        .Type = adTypeText
        .Mode = adModeReadWrite
        .Charset = Charset
        .Open
        
        For i = 0 To UBound(vData)
            vRow = vData(i)
            For j = 0 To UBound(vRow) - 1
                stream = stream & vRow(j) & SplitChar
            Next
            stream = stream & vRow(j) & LineChar
            
            If i Mod 10 = 0 Then
                .WriteText stream
                stream = ""
            End If
        Next i
        
        If Len(stream) > 0 Then
            .WriteText stream
        End If
        
        .SaveToFile locDir, adSaveCreateNotExist
        .Flush
        .Close
    End With
    
    Set oStream = Nothing
End Function

Public Sub IndexFirst()
    ThisPoint = 0
End Sub

Public Function IndexPrevious() As Boolean
    ThisPoint = ThisPoint - 1
    If ThisPoint > -1 Then
        IndexPrevious = True
    Else
        IndexPrevious = False
    End If
End Function

Public Function IndexNext() As Boolean
    ThisPoint = ThisPoint + 1
    If ThisIsGroup Then
        If ThisIndexCount > ThisPoint Then
            IndexNext = True
        Else
            IndexNext = False
        End If
    Else
        If ThisCount > ThisPoint Then
            IndexNext = True
        Else
            IndexNext = False
        End If
    End If
End Function

Public Sub IndexLast()
    If ThisIsGroup Then
        ThisPoint = ThisIndexCount - 1
    Else
        ThisPoint = ThisCount - 1
    End If
End Sub

Public Function GetDataByIndex() As Variant
    Dim l As Long, lCount As Long
    Dim vCol() As String
    Dim vData As Variant
    Dim s As String
    
    If ThisIsGroup Then
        vCol = Split(ThisIndexValue(ThisPoint), ".")
        lCount = UBound(vCol)
        ReDim vData(lCount)
        
        For l = 0 To lCount
            vData(l) = ThisData(vCol(l))
        Next l
        Erase vCol
    Else
        vData = ThisData(ThisPoint)
    End If
    
    GetDataByIndex = vData
    Erase vData
End Function

Public Function GetKeyByIndex() As String
    If ThisIsGroup Then
        GetKeyByIndex = ThisIndexKey(ThisPoint)
    End If
End Function

Public Function LoadAs(ByVal locDir As String, Optional ByVal Charset As String = "UTF-8") As String
    Dim oStream As ADODB.stream
    Set oStream = New ADODB.stream
    
    With oStream
        .Type = adTypeText
        .Mode = adModeReadWrite
        .Charset = Charset
        .Open
        .LoadFromFile locDir
        LoadAs = .ReadText
        .Close
    End With
    
    Set oStream = Nothing
End Function

Public Function SaveAs(ByVal locDir As String, stream As String, Optional ByVal Charset As String = "UTF-8") As Boolean
    Dim oStream As ADODB.stream
    Set oStream = New ADODB.stream
    
    If Len(Dir(locDir, vbArchive)) > 0 Then
        FileSystem.Kill locDir
    End If
    
    With oStream
        .Type = adTypeText
        .Mode = adModeReadWrite
        .Charset = Charset
        .Open
        .WriteText stream
        .SaveToFile locDir, adSaveCreateNotExist
        .Flush
        .Close
    End With
    
    Set oStream = Nothing
End Function

Private Sub Dispose()
    ThisPoint = -1
    ThisCount = 0
    ThisIndexCount = 0
    Erase ThisIndexKey
    Erase ThisIndexValue
    Erase ThisData
    
    If ThisIsGroup Then
        Erase ThisIndexs
        ThisIsGroup = False
    End If
End Sub
