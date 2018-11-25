VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LocalContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetRemotes() As TBRemotes
    Dim Index As Long, localDir As String
    Dim vData() As String, vLine() As String
    Dim rmts As TBRemotes, rmt As TYPE_REMOTE
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_REMOTE
    
    Set rmts = New TBRemotes
    vData = GetData(localDir)
    
    For Index = 0 To UBound(vData)
        vLine = Split(vData(Index), vbTab)
        With rmt
            .RootID = vLine(0)
            .name = vLine(1)
            .RangeID = vLine(2)
            .DownRoot = vLine(3)
            .UpRoot = vLine(4)
            .CacheRoot = vLine(5)
            .BackupRoot = vLine(6)
            .Server = vLine(7)
            .Username = vLine(8)
            .Password = vLine(9)
        End With
        
        rmts.Add rmt
        Erase vLine
    Next Index
    
    Set GetRemotes = rmts
    Set rmts = Nothing
    Erase vData
End Function

Public Sub SetRemotes(rmts As TBRemotes)
    Dim Index As Long, localDir As String
    Dim rmt As TYPE_REMOTE
    Dim vData() As String, sRow As String
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_REMOTE
    ReDim vData(rmts.Size - 1)
    
    For Index = 0 To rmts.Size - 1
        rmt = rmts.Remote(Index)
        With rmt
            sRow = .RootID & vbTab & .name & vbTab & .RangeID & vbTab & .DownRoot & vbTab & .UpRoot & vbTab & .CacheRoot & vbTab & .BackupRoot _
            & vbTab & .Server & vbTab & .Username & vbTab & .Password
        End With
        vData(Index) = sRow
    Next
    SetData localDir, vData
    Erase vData
End Sub

Public Function GetMetadatas() As TBMetadatas
    Dim Index As Long, localDir As String
    Dim vData() As String, vLine() As String
    Dim metas As TBMetadatas, meta As TYPE_METADATA
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_METADATA
    Set metas = New TBMetadatas
    vData = GetData(localDir)
    
    For Index = 0 To UBound(vData)
        vLine = Split(vData(Index), vbTab)
        With meta
            .ID = vLine(0)
            .FromID = vLine(1)
            .ToID = vLine(2)
            .FileName = vLine(3)
            .DtFormat = vLine(4)
            .SnFormat = vLine(5)
            .Postfix = vLine(6)
            .Directory = vLine(7)
        End With
        
        metas.Add meta
        Erase vLine
    Next Index
    
    Set GetMetadatas = metas
    Set metas = Nothing
    Erase vData
End Function

Public Function GetK3Server() As TYPE_K3SERVER
    Dim Section As String, localDir As String
    Dim svr As TYPE_K3SERVER
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_K3SERVER
    mINI.Config_INI_Patch = localDir
    
    Section = "DBConnection"
    svr.DBServer = mINI.GetIniValue(Section, "DB_Server", ".")
    svr.DBName = mINI.GetIniValue(Section, "DB_Database", "")
    svr.DBUsername = mINI.GetIniValue(Section, "DB_Username", "sa")
    svr.DBPassword = mINI.GetIniValue(Section, "DB_Password", "")
    
    Section = "K3"
    svr.K3Username = mINI.GetIniValue(Section, "K3_Username", "administrator")
    
    GetK3Server = svr
End Function

Public Sub SetK3Server(svr As TYPE_K3SERVER)
    Dim Section As String, localDir As String
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_K3SERVER
    mINI.Config_INI_Patch = localDir
    
    Section = "DBConnection"
    mINI.SetIniValue Section, "DB_Server", svr.DBServer
    mINI.SetIniValue Section, "DB_Database", svr.DBName
    mINI.SetIniValue Section, "DB_Username", svr.DBUsername
    mINI.SetIniValue Section, "DB_Password", svr.DBPassword

    Section = "K3"
    mINI.SetIniValue Section, "K3_Username", svr.K3Username
End Sub

Public Function GetTasks() As TB_Context.TBTasks
    Dim Index As Long, localDir As String
    Dim vData() As String, vLine() As String
    Dim Tasks As TB_Context.TBTasks, task As TYPE_TASK
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_TASK
    
    Set Tasks = New TB_Context.TBTasks
    vData = GetData(localDir)
    
    For Index = 0 To UBound(vData)
        vLine = Split(vData(Index), vbTab)
        With task
            .Number = vLine(0)
            .IsSys = IIf(vLine(1) = "1", True, False)
            .IsUse = IIf(vLine(2) = "1", True, False)
            .EXEName = vLine(3)
            .Parameter = vLine(4)
            .RunStyle = vLine(5)
            .Interval = vLine(6)
            .StartTime = vLine(7)
            .StopTime = vLine(8)
            .Description = vLine(9)
        End With
        
        If IsRightTasks(task.Description) Then
            Tasks.Add task
        End If
        Erase vLine
    Next Index
    
    Set GetTasks = Tasks
    Set Tasks = Nothing
    Erase vData
End Function

Private Function IsRightTasks(ByVal name As String) As Boolean
'    Select Case name
'        Case "Download files"
'            IsRightTasks = True
'        Case "Upload files"
'            IsRightTasks = True
'        Case "Product and Price interface"
'            IsRightTasks = True
'        Case "Shipping information"
'            IsRightTasks = True
'        Case "Sales Transaction"
'            IsRightTasks = True
'        Case "Stock Movement(W-S)"
'            IsRightTasks = True
'        Case "Stock Movement(S-W)"
'            IsRightTasks = True
'        Case "Stock Movement(B-B)"
'            IsRightTasks = True
'        Case "Stock Movement(S-S)"
'            IsRightTasks = True
'        Case "Stock Movement(W-W)"
'            IsRightTasks = True
'        Case "Stock Image Report"
'            IsRightTasks = True
'        Case "Send log"
'            IsRightTasks = True
'    End Select
IsRightTasks = True
End Function

Public Sub SetTasks(Tasks As TB_Context.TBTasks)
    Dim Index As Long, localDir As String
    Dim task As TYPE_TASK
    Dim vData() As String, sRow As String
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_TASK
    ReDim vData(Tasks.Size - 1)
    
    For Index = 0 To Tasks.Size - 1
        task = Tasks.task(Index)
        With task
            sRow = .Number & vbTab & IIf(.IsSys, "1", "0") & vbTab & IIf(.IsUse, "1", "0") & vbTab & .EXEName & vbTab & .Parameter & vbTab _
                & .RunStyle & vbTab & .Interval & vbTab & .StartTime & vbTab & .StopTime & vbTab & .Description
        End With
        vData(Index) = sRow
    Next
    SetData localDir, vData
    Erase vData
End Sub

Public Function GetActions() As TB_Context.TBActions
    Dim Index As Long, localDir As String
    Dim vData() As String, vLine() As String
    Dim acts As TB_Context.TBActions, act As TYPE_ACTION
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_ACTION
    
    Set acts = New TB_Context.TBActions
    vData = GetData(localDir)
    
    For Index = 0 To UBound(vData)
        vLine = Split(vData(Index), vbTab)
        With act
            .TaskNumber = vLine(0)
            .ActionNumber = vLine(1)
            .MetadataNumber = vLine(2)
            .DllName = vLine(3)
        End With
        
        acts.Add act
        Erase vLine
    Next Index
    
    Set GetActions = acts
    Set acts = Nothing
    Erase vData
End Function

Public Function GetSmtp() As TYPE_EMAILSMTP
    Dim Section As String, localDir As String
    Dim smp As TYPE_EMAILSMTP
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_SMTP
    mINI.Config_INI_Patch = localDir
    
    Section = "SMTP"
    smp.Smtp = mINI.GetIniValue(Section, "SMTP_Server", "")
    smp.Username = mINI.GetIniValue(Section, "SMTP_Username", "")
    smp.Password = mINI.GetIniValue(Section, "SMTP_Password", "")
    smp.Port = CInt(Val(mINI.GetIniValue(Section, "SMTP_Port", "25")))
    smp.Sender = mINI.GetIniValue(Section, "SMTP_Sender", "")
    
    GetSmtp = smp
End Function

Public Function SetSmtp(smp As TYPE_EMAILSMTP)
    Dim Section As String, localDir As String
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_SMTP
    mINI.Config_INI_Patch = localDir
    
    Section = "SMTP"
    mINI.SetIniValue Section, "SMTP_Server", smp.Smtp
    mINI.SetIniValue Section, "SMTP_Username", smp.Username
    mINI.SetIniValue Section, "SMTP_Password", smp.Password
    mINI.SetIniValue Section, "SMTP_Port", smp.Port
    mINI.SetIniValue Section, "SMTP_Sender", smp.Sender
End Function

Public Function GetMails() As TB_Context.TBMailsEx
    Dim Index As Long, l As Long, localDir As String
    Dim vData() As String, vLine() As String
    Dim vMail() As String
    Dim eTo As TB_Context.TBMails, eCC As TB_Context.TBMails
    Dim eMails As TB_Context.TBMailsEx
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_EMAIL
    Set eMails = New TB_Context.TBMailsEx
    
    vData = GetData(localDir)
    For Index = 0 To UBound(vData)
        vLine = Split(vData(Index), vbTab)
        
        vMail = Split(vLine(1), ";")
        Set eTo = New TB_Context.TBMails
        For l = 0 To UboundEx(vMail)
            If Len(vMail(l)) > 0 Then
                eTo.Add vMail(l)
            End If
        Next l
        Erase vMail
        
        vMail = Split(vLine(2), ";")
        Set eCC = New TB_Context.TBMails
        For l = 0 To UboundEx(vMail)
            If Len(vMail(l)) > 0 Then
                eCC.Add vMail(l)
            End If
        Next l
        Erase vMail
        
        eMails.AddToMail vLine(0), eTo
        eMails.AddCCMail vLine(0), eCC
        
        Set eTo = Nothing
        Set eCC = Nothing
        Erase vLine
    Next Index
    
    Set GetMails = eMails
    Set eMails = Nothing
    
    Erase vData
End Function

Public Sub SetMails(mails As TB_Context.TBMailsEx)
    Dim Index As Long, l As Long, lCount As Long, localDir As String
    Dim vData() As String
    Dim eTo As TB_Context.TBMails, eCC As TB_Context.TBMails
    Dim eMails As TB_Context.TBMailsEx
    
    Dim sData As String, sRow As String
    
    localDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
    localDir = localDir & CONST_DIR_SETTING & CONST_FILE_EMAIL
    
    For Index = 0 To mails.Size - 1
        
        sRow = mails.Key(Index) & vbTab
        
        Set eTo = mails.ToMail(Index)
        If Not eTo Is Nothing Then
            If eTo.Size > 0 Then
                For l = 0 To eTo.Size - 1
                    If Len(eTo.Mail(l)) > 0 Then
                        sRow = sRow & eTo.Mail(l) & ";"
                    End If
                Next l
                
                sRow = Left(sRow, Len(sRow) - 1)
            End If
        End If
        sRow = sRow & vbTab
        Set eTo = Nothing
        
        Set eCC = mails.CCMail(Index)
        If Not eCC Is Nothing Then
            If eCC.Size > 0 Then
                For l = 0 To eCC.Size - 1
                    If Len(eCC.Mail(l)) > 0 Then
                        sRow = sRow & eCC.Mail(l) & ";"
                    End If
                Next l
                
                sRow = Left(sRow, Len(sRow) - 1)
            End If
        End If
        Set eTo = Nothing
        
        ReDim Preserve vData(lCount)
        vData(lCount) = sRow
        lCount = lCount + 1
    Next Index
    
    SetData localDir, vData
    Erase vData
End Sub

Private Function GetData(localDir As String) As String()
    Dim Index As Long, Point As Long
    Dim vRow() As String, sRow As String
    
    Point = FreeFile
    Open localDir For Input As #Point
    While Not EOF(Point)
        Line Input #Point, sRow
        sRow = Trim(sRow)
        If Len(sRow) > 0 Then
            If Left(sRow, 1) <> "[" Then
                ReDim Preserve vRow(Index)
                vRow(Index) = sRow
                Index = Index + 1
            End If
        End If
    Wend
    Close #Point
    
    GetData = vRow
    Erase vRow
End Function

Private Function SetData(localDir As String, vData() As String)
    Dim Index As Long, Size As Long, Point As Long
    
    Point = FreeFile
    Size = UBound(vData)
    Open localDir For Output As #Point
    For Index = 0 To Size
        Print #Point, vData(Index)
    Next Index
    Close #Point
End Function

Private Function UboundEx(v As Variant) As Long
    On Error GoTo HERROR
    UboundEx = UBound(v)
    Exit Function
HERROR:
    UboundEx = -1
End Function
