VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XZShareLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CONST_IDENTITY_LEN As Integer = 14
Private Const CONST_IDENTITY_TIMEOUT As Integer = 60    'min

Private ThisMemoryIndex As Long
Private ThisMemory As XZMemory

Private Sub Class_Initialize()
    InitShare
End Sub

Private Sub Class_Terminate()
    If ThisMemoryIndex > 0 Then
        SetShareAsNothing
        ThisMemoryIndex = 0
    End If
    Set ThisMemory = Nothing
End Sub

Private Sub InitShare()
    Dim lIndex As Long
    Dim sha As TYPE_XZ_SHARE
    
    Set ThisMemory = New XZMemory
    
    If ThisMemory.NewMemory Then
        For lIndex = 1 To ThisMemory.Count
            sha = GetShareByIndex(lIndex)
            If IsLostEfficacy(sha.Identity) Then
                ThisMemoryIndex = lIndex
                SetShareAsNullString
                Exit For
            End If
        Next lIndex
        
        If ThisMemoryIndex = 0 Then
            Err.Raise -1, "", "没有空闲的共享内存通道!"
        End If
    Else
        Err.Raise -1, "", "无法映射内存!"
    End If
End Sub

Private Function GetIdentiryByIndex(Index As Long) As String
    Dim res As String
    res = ThisMemory.InputMemory(Index)
    GetIdentiryByIndex = IIf(Len(res) >= CONST_IDENTITY_LEN, Left(res, CONST_IDENTITY_LEN), String(CONST_IDENTITY_LEN, " "))
End Function

Private Function IsLostEfficacy(Identity As String) As Boolean
    Dim vNow(5) As Integer
    Dim vNew(1) As String
    
    IsLostEfficacy = True
    
    If IsNumeric(Identity) And Len(Identity) = CONST_IDENTITY_LEN Then
        vNow(0) = CInt(Mid(Identity, 1, 4))
        vNow(1) = CInt(Mid(Identity, 5, 2))
        vNow(2) = CInt(Mid(Identity, 7, 2))
        vNow(3) = CInt(Mid(Identity, 9, 2))
        vNow(4) = CInt(Mid(Identity, 11, 2))
        vNow(5) = CInt(Mid(Identity, 13, 2))
        
        vNew(0) = vNow(0) & "-" & vNow(1) & "-" & vNow(2) & " " & vNow(3) & ":" & vNow(4) & ":" & vNow(5)
        vNew(1) = Format(Now, "YYYY-MM-DD hh:mm:ss")
        
        If Abs(DateDiff("n", vNew(0), vNew(1))) <= CONST_IDENTITY_TIMEOUT Then
            IsLostEfficacy = False
        End If
    End If
    
    Erase vNow
    Erase vNew
End Function

Private Function GetShareByIndex(Index As Long) As TYPE_XZ_SHARE
    Dim Notice As String
    Dim sha As TYPE_XZ_SHARE
    
    Notice = ThisMemory.InputMemory(Index)
    If Len(Notice) >= CONST_IDENTITY_LEN Then
        sha.Identity = Left(Notice, CONST_IDENTITY_LEN)
        sha.Resource = Right(Notice, Len(Notice) - CONST_IDENTITY_LEN)
    End If
    GetShareByIndex = sha
End Function

Private Function GetShareByNotice(Notice As String) As TYPE_XZ_SHARE
    Dim sha As TYPE_XZ_SHARE
    If Len(Notice) >= CONST_IDENTITY_LEN Then
        sha.Identity = Left(Notice, CONST_IDENTITY_LEN)
        sha.Resource = Right(Notice, Len(Notice) - CONST_IDENTITY_LEN)
    End If
    GetShareByNotice = sha
End Function

Private Function SetShareAsNothing() As Boolean
    ThisMemory.OutputMemory ThisMemoryIndex, ""
    SetShareAsNothing = True
End Function

Private Function SetShareAsNullString() As Boolean
    SetShareByResouce ""
    SetShareAsNullString = True
End Function

Private Function SetShareByResouce(Resource As String) As TYPE_XZ_SHARE
    Dim sha As TYPE_XZ_SHARE
    sha.Identity = Format(Now, "YYYYMMDDhhmmss")
    sha.Resource = Resource
    ThisMemory.OutputMemory ThisMemoryIndex, sha.Identity & sha.Resource
    SetShareByResouce = sha
End Function

Private Function IsLockByResource(Resource As String) As Boolean
    Dim Index As Long
    Dim sha As TYPE_XZ_SHARE
    
    For Index = 1 To ThisMemory.Count
        If Index <> ThisMemoryIndex Then
            sha = GetShareByIndex(Index)
            If sha.Resource = Resource And Not IsLostEfficacy(sha.Identity) Then
                IsLockByResource = True
                Exit Function
            End If
        End If
    Next Index
End Function

Public Sub DoLock(sha As TYPE_XZ_SHARE)
    Dim shaNew As TYPE_XZ_SHARE
    
    shaNew = GetShareByIndex(ThisMemoryIndex)
    If Not IsLostEfficacy(shaNew.Identity) And Len(shaNew.Resource) > 0 Then
        Err.Raise -1, "", "当前内存共享通道已锁定其他的资源!"
    End If
    shaNew = SetShareByResouce(sha.Resource)
    If Len(sha.Resource) > 0 And IsLockByResource(shaNew.Resource) Then
        SetShareAsNullString
        Err.Raise -1, "", "当前资源已被锁定!"
    End If
    sha.Identity = shaNew.Identity
End Sub

Public Function IsLock(sha As TYPE_XZ_SHARE) As Boolean
    Dim shaNew As TYPE_XZ_SHARE
    
    shaNew = GetShareByIndex(ThisMemoryIndex)
    If Not IsLostEfficacy(shaNew.Identity) And Len(shaNew.Resource) > 0 And shaNew.Resource = sha.Resource Then
        IsLock = True
        Exit Function
    End If
    If Len(sha.Resource) > 0 And IsLockByResource(sha.Resource) Then
        IsLock = True
        Exit Function
    End If
End Function

Public Sub DeLock(sha As TYPE_XZ_SHARE)
    Dim shaNew As TYPE_XZ_SHARE
    
    shaNew = GetShareByIndex(ThisMemoryIndex)
    If IsLostEfficacy(shaNew.Identity) Or (shaNew.Identity = sha.Identity And shaNew.Resource = sha.Resource) Then
        SetShareAsNullString
    Else
        Err.Raise -1, "", "当前资源与锁定的资源不一致!"
    End If
    sha.Identity = ""
End Sub



