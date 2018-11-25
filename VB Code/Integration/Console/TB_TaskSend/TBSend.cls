VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ILoad

Private ThisK3Svr As TYPE_K3SERVER
Private ThisSmtp As TYPE_EMAILSMTP
Private ThisSends As TB_Context.TBMailsEx
Private ThisTasks As TB_Context.TBTasks
Private ThisParameter As TB_Context.TBParameters

Private ThisSetCaption() As String
Private ThisSetField() As String
Private ThisSetWidth() As String
Private ThisSetHeight As String
Private ThisSetBackground As String
Private ThisSetBackgroundColor As String
Private ThisSetBorderColor As String
Private ThisSetBorderWidth As String
Private ThisSetBorderStyle As String
Private ThisSetFont As String
Private ThisSetFontSize As String

Private Const CONST_TABLENAME_LOG As String = "t_TB_Log"
Private Const CONST_FILEDIR_LOG As String = "Log\"
Private Const CONST_FILEDIR_SET As String = "Setting\"
Private Const CONST_FILENAME_SET As String = "LogReport.ini"

Private Function ILoad_Init(param As TB_Context.TBParameters) As Boolean
On Error GoTo HERROR
    
    Set ThisParameter = param
    
    LogEx "Log Report Initialize"
    DoEvents
    
    Dim o As TB_Runtime.LocalContent
    Set o = New TB_Runtime.LocalContent
    ThisK3Svr = o.GetK3Server
    ThisSmtp = o.GetSmtp
    Set ThisSends = o.GetMails
    Set ThisTasks = o.GetTasks
    Set o = Nothing
    
    Dim locDir As String
    Dim Section As String
    Section = "LOG"
    locDir = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    locDir = locDir & CONST_FILEDIR_SET & CONST_FILENAME_SET
    mINI.Config_INI_Patch = locDir
    ThisSetCaption = Split(mINI.GetIniValue(Section, "CAPTION", "NO.;Action;Date;From;To;File Name;Description"), ";")
    ThisSetField = Split(mINI.GetIniValue(Section, "FIELD", "FID;FAction;FDate;FFrom;FTo;FFileName;FDescription"), ";")
    ThisSetWidth = Split(mINI.GetIniValue(Section, "WIDTH", "50;50;120;50;50;180;500"), ";")
    ThisSetHeight = mINI.GetIniValue(Section, "HEIGHT", "20")
    ThisSetBackground = mINI.GetIniValue(Section, "BACKGROUND", "silver")
    ThisSetBackgroundColor = mINI.GetIniValue(Section, "BACKGROUND COLOR", "transparent")
    ThisSetBorderColor = mINI.GetIniValue(Section, "BORDER COLOR", "black")
    ThisSetBorderWidth = mINI.GetIniValue(Section, "BORDER WIDTH", "0.5pt")
    ThisSetBorderStyle = mINI.GetIniValue(Section, "BORDER STYLE", "solid")
    ThisSetFont = mINI.GetIniValue(Section, "FONT", "Times New Roman")
    ThisSetFontSize = mINI.GetIniValue(Section, "FONT SIZE", "14px")
    
    ILoad_Init = True
    Exit Function

HERROR:
    LogEx Err.Description
End Function

Private Sub ILoad_Run(param As TB_Context.TBParameters)
    LogEx "Log Report Running"
    DoEvents
    DoAction
End Sub

Private Sub ILoad_Dispose(param As TB_Context.TBParameters)
    
    LogEx "Log Report Dispose"
    DoEvents
    
    Set ThisSends = Nothing
    Set ThisTasks = Nothing
    Set ThisParameter = Nothing
    
    Erase ThisSetCaption
    Erase ThisSetField
    Erase ThisSetWidth
End Sub

Private Sub DoAction()
    Dim Index As Long
    
    For Index = 0 To ThisSends.Size - 1
        Send Index
        DoEvents
    Next Index
End Sub

Private Sub Send(Index As Long)
    Dim Index2 As Long
    Dim TaskNumber As String, TaskDesc As String
    Dim locDir As String, sHml As String
    Dim rs As ADODB.Recordset
On Error GoTo HERROR
    
    TaskNumber = ThisSends.Key(Index)
    Index2 = ThisTasks.Lookup(TaskNumber)
    TaskDesc = ThisTasks.Task(Index2).Description
    
    LogEx "send '" & TaskDesc & "'"
    
    Set rs = GetLogs(TaskNumber)
    
    If rs.RecordCount > 0 Then
        sHml = OutputHtmls(TaskNumber, TaskDesc, rs, locDir)
        SendMailEx locDir, sHml, TaskNumber, TaskDesc
        SetSend rs
        
        LogEx "send '" & locDir & "' success!"
    Else
        LogEx "There is no need to send to the data."
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Sub
HERROR:
    LogEx "send failed! " & Err.Description
    Set rs = Nothing
End Sub


Private Function GetLogs(ByVal TaskNumber As String) As ADODB.Recordset
    Dim sSQL As String
    sSQL = "SELECT * FROM " & CONST_TABLENAME_LOG & " WHERE FIsSend = 1 AND FTaskNumber = '" & TaskNumber & "' ORDER BY FID"
    Set GetLogs = TB_Runtime.ExecuteSQL(sSQL)
End Function

'Private Function OutputHtmls(rs As ADODB.Recordset) As String
'    Dim locDir As String, locName As String
'    Dim lPt As Long, Index As Long, Index2 As Long, lCount As Long
'    Dim sHml As String
'
'    locDir = IIf(Right(App.path, 1) = "\", App.path, App.path & "\")
'    locDir = locDir & CONST_FILEDIR_LOG & "TASK" & mStart.Parameter.Value("@TASK") & "\"
'    MakeDirByLocal locDir
'    locName = Format(Now, "YYYYMMDDhhmmss") & ".htm"
'    lCount = UboundEx(ThisSetCaption)
'
'    lPt = FreeFile
'    Open locDir & locName For Output As #lPt
'
'    sHml = "<HTML><HEAD><TITLE>By Fun4.org</TITLE></HEAD><META http-equiv=Content-Type content=""text/html; charset=gb2312""><BODY>" & vbCrLf
'    sHml = sHml & "<DIV><TABLE STYLE=""border-collapse:collapse; border:0; cellSpacing:0; cellPadding:0; word-break:break-all; word-wrap:break-all""><TR height=" & ThisSetHeight(0) & ">" & vbCrLf
'    sHml = sHml & "<!-- By JianJian.org -->" & vbCrLf
'
'    For Index = 0 To lCount
'        sHml = sHml & "    <TD style=""border-width:" & ThisSetBorderWidth & "; border-style:" & ThisSetBorderStyle & "; border-color:" & ThisSetBorderColor & "; background-color:" & ThisSetBackgroundColor & "; width:" & ThisSetWidth(Index) & "; background:" & ThisSetBackground & """>" & vbCrLf
'        sHml = sHml & "    <P align=""center""><STRONG><FONT size=""" & ThisSetFontSize & """ face=""" & ThisSetFont & """>" & ThisSetCaption(Index) & "</FONT></STRONG></P></TD>" & vbCrLf
'    Next Index
'    sHml = sHml & "</TR>"
'
'    Print #lPt, sHml: sHml = ""
'
'    For Index = 0 To rs.RecordCount - 1
'        sHml = "<TR height=" & ThisSetHeight(1) & ">" & vbCrLf
'        For Index2 = 0 To lCount
'            sHml = sHml & "        <TD style=""border-width:" & ThisSetBorderWidth & "; border-style:" & ThisSetBorderStyle & "; border-color:" & ThisSetBorderColor & "; background-color:" & ThisSetBackgroundColor & "; width:" & ThisSetWidth(Index2) & """>" & vbCrLf
'            sHml = sHml & "        <FONT size=""" & ThisSetFontSize & """ face=""" & ThisSetFont & """>" & rs(ThisSetField(Index2)) & "</FONT></TD>" & vbCrLf
'        Next Index2
'        sHml = sHml & "</TR>"
'        Print #lPt, sHml: sHml = ""
'        rs.MoveNext
'    Next Index
'
'    sHml = "</TABLE></DIV></BODY></HTML>"
'    Print #lPt, sHml: sHml = ""
'    Close #lPt
'    OutputHtmls = locDir & locName
'End Function

Private Function OutputHtmls(TaskName As String, TaskDesc As String, rs As ADODB.Recordset, locDir As String) As String
    Dim locName As String
    Dim lPt As Long, Index As Long, Index2 As Long, lCount As Long
    Dim sHml As String, sHml2 As String
    
    locDir = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    locDir = locDir & CONST_FILEDIR_LOG & "TASK" & ThisParameter.Value("@TASK") & "\"
    MakeDirByLocal locDir
    locName = Format(Now, "YYYYMMDDhhmmss") & ".htm"
    lCount = UboundEx(ThisSetCaption)
    
    lPt = FreeFile
    Open locDir & locName For Output As #lPt
    
    sHml = "<HTML><HEAD><!-- By Fun4.org -->" & vbCrLf
    sHml = sHml & "<TITLE>Log Report</TITLE><STYLE type=""text/css"">" & vbCrLf
    sHml = sHml & "    * {font-family:" & ThisSetFont & "; font-size:" & ThisSetFontSize & ";}" & vbCrLf
    sHml = sHml & "    TABLE {width:100%; border-collapse:collapse; border:0; cellSpacing:0; cellPadding:0; word-break:break-all; word-wrap:break-all;}" & vbCrLf
    sHml = sHml & "    TR {height:" & ThisSetHeight & ";}" & vbCrLf
    sHml = sHml & "    TD {border-width:" & ThisSetBorderWidth & "; border-style:" & ThisSetBorderStyle & "; border-color:" & ThisSetBorderColor & "; background-color:" & ThisSetBackgroundColor & ";}" & vbCrLf
    sHml = sHml & "    P {text-align:center; font-weight:bold;}" & vbCrLf
    sHml = sHml & "</STYLE></HEAD><BODY>" & vbCrLf
    
    sHml = sHml & "<DIV>" & vbCrLf
    sHml = sHml & "    <TABLE>" & vbCrLf
    sHml = sHml & "    <TR><TD style=""width:200; background:" & ThisSetBackground & """><P>Interface Name</P></TD><TD>" & TaskDesc & "</TD></TR>" & vbCrLf
    sHml = sHml & "    <TR><TD style=""width:200; background:" & ThisSetBackground & """><P>Date</P></TD><TD>" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "</TD></TR>" & vbCrLf
    sHml = sHml & "    <TR><TD style=""width:200; background:" & ThisSetBackground & """><P>Log Number</P></TD><TD>" & rs.RecordCount & "</TD></TR>" & vbCrLf
    sHml = sHml & "    </TABLE>" & vbCrLf
    sHml = sHml & "</DIV>"
    
    Print #lPt, sHml: sHml = ""
    
    sHml = "<DIV style=""margin-top:20px""><TABLE>" & vbCrLf
    sHml = sHml & "<TR>" & vbCrLf
    
    For Index = 0 To lCount
        sHml = sHml & "    <TD style=""width:" & ThisSetWidth(Index) & ";background:" & ThisSetBackground & """><P>" & ThisSetCaption(Index) & "</P></TD>" & vbCrLf
    Next Index
    sHml = sHml & "</TR>"
    
    Print #lPt, sHml: sHml = ""
    
    For Index = 0 To rs.RecordCount - 1
        sHml = "<TR>"
        For Index2 = 0 To lCount
            sHml = sHml & "<TD>" & rs(ThisSetField(Index2)) & "</TD>"
        Next Index2
        sHml = sHml & "</TR>"
        Print #lPt, sHml: sHml = ""
        rs.MoveNext
    Next Index
    
    sHml = "</TABLE></DIV></BODY></HTML>"
    
    sHml2 = "<HTML><HEAD><!-- By JianJian.org --><TITLE>Log Report</TITLE></HEAD><BODY>" & vbCrLf
    sHml2 = sHml2 & "<DIV>" & vbCrLf
    sHml2 = sHml2 & "   <TABLE  style=""border-collapse:collapse; border:0; cellSpacing:0; cellPadding:0; word-break:break-all; word-wrap:break-all"">" & vbCrLf
    sHml2 = sHml2 & "   <TR height=20>" & vbCrLf
    sHml2 = sHml2 & "       <TD style=""border-width:" & ThisSetBorderWidth & "; border-style:" & ThisSetBorderStyle & "; border-color:" & ThisSetBorderColor & "; background-color:" & ThisSetBackgroundColor & "; width:200; background:silver"">" & vbCrLf
    sHml2 = sHml2 & "       <P align=""center""><STRONG><FONT size=""2"" face=""" & ThisSetFont & """>Interface Name</FONT></STRONG></P></TD>" & vbCrLf
    sHml2 = sHml2 & "       <TD style=""border-width:" & ThisSetBorderWidth & "; border-style:" & ThisSetBorderStyle & "; border-color:" & ThisSetBorderColor & "; background-color:" & ThisSetBackgroundColor & "; width:300;"">" & vbCrLf
    sHml2 = sHml2 & "       <FONT size=""2"" face=""" & ThisSetFont & """>" & TaskDesc & "</FONT></TD>" & vbCrLf
    sHml2 = sHml2 & "   </TR><TR height=20>" & vbCrLf
    sHml2 = sHml2 & "       <TD style=""border-width:" & ThisSetBorderWidth & "; border-style:" & ThisSetBorderStyle & "; border-color:" & ThisSetBorderColor & "; background-color:" & ThisSetBackgroundColor & "; width:200; background:silver"">" & vbCrLf
    sHml2 = sHml2 & "       <P align=""center""><STRONG><FONT size=""2"" face=""" & ThisSetFont & """>Date</FONT></STRONG></P></TD>" & vbCrLf
    sHml2 = sHml2 & "       <TD style=""border-width:" & ThisSetBorderWidth & "; border-style:" & ThisSetBorderStyle & "; border-color:" & ThisSetBorderColor & "; background-color:" & ThisSetBackgroundColor & "; width:300;"">" & vbCrLf
    sHml2 = sHml2 & "       <FONT size=""2"" face=""" & ThisSetFont & """>" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "</FONT></TD>" & vbCrLf
    sHml2 = sHml2 & "   </TR><TR height=20>" & vbCrLf
    sHml2 = sHml2 & "       <TD style=""border-width:" & ThisSetBorderWidth & "; border-style:" & ThisSetBorderStyle & "; border-color:" & ThisSetBorderColor & "; background-color:" & ThisSetBackgroundColor & "; width:200; background:silver"">" & vbCrLf
    sHml2 = sHml2 & "       <P align=""center""><STRONG><FONT size=""2"" face=""" & ThisSetFont & """>Log Number</FONT></STRONG></P></TD>" & vbCrLf
    sHml2 = sHml2 & "       <TD style=""border-width:" & ThisSetBorderWidth & "; border-style:" & ThisSetBorderStyle & "; border-color:" & ThisSetBorderColor & "; background-color:" & ThisSetBackgroundColor & "; width:300;"">" & vbCrLf
    sHml2 = sHml2 & "       <FONT size=""2"" face=""" & ThisSetFont & """>" & rs.RecordCount & "</FONT></TD>" & vbCrLf
    sHml2 = sHml2 & "   </TR></TABLE></DIV>" & vbCrLf
    sHml2 = sHml2 & "   <P><STRONG><FONT size=""2"" face=""" & ThisSetFont & """>Please check the attachment!</FONT></STRONG></P>" & vbCrLf
    sHml2 = sHml2 & "</BODY></HTML>"
    
    Print #lPt, sHml: sHml = ""
    Close #lPt
    
'    Dim rar As String
'    rar = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "Rar.exe a """ & CONST_FILEDIR_LOG & "TASK" & mStart.Parameter.Value("@TASK") & "\" & locName & """ """ & CONST_FILEDIR_LOG & "TASK" & mStart.Parameter.Value("@TASK") & "\" & locName & ".rar"""
'    LogEx rar
'    Shell rar
    
    locDir = locDir & locName
    OutputHtmls = sHml2
End Function

Private Sub SendMail(locDir As String, sHml As String, TaskNumber As String, TaskDesc As String)
    Dim oMail As Object
    Dim oMails As TB_Context.TBMails
    Dim Index As Long, Index2 As Long
    
    Index = ThisSends.Lookup(TaskNumber)
    
    If ThisSends.ToMail(Index).Size > 0 Or ThisSends.CCMail(Index).Size > 0 Then
        Set oMail = CreateObject("jmail.Message")
        
        With oMail
            .Charset = "gb2312"
            .Silent = False
            '.ContentType = "text/html"
            .Priority = 1
            
            .MailServerUserName = ThisSmtp.Username
            .MailServerPassWord = ThisSmtp.Password
            
            .From = ThisSmtp.Sender
            
            Set oMails = ThisSends.ToMail(Index)
            For Index2 = 0 To oMails.Size - 1
                .AddRecipient oMails.Mail(Index2)
            Next Index2
            Set oMails = Nothing
            
            Set oMails = ThisSends.CCMail(Index)
            For Index2 = 0 To oMails.Size - 1
                .AddRecipientCC oMails.Mail(Index2)
            Next Index2
            Set oMails = Nothing
            
            .AddAttachment locDir
            .Subject = TaskDesc & " -Log Report " & Format(Now, "YYYY/MM/DD hh:mm:ss")
            
            .Htmlbody = sHml
            .Body = "Please check the attachment!"
            
            .Send ThisSmtp.Smtp
        End With
        
        Set oMail = Nothing
    End If
    
    Set oMails = Nothing
End Sub

Private Sub SendMailEx(locDir As String, sHml As String, TaskNumber As String, TaskDesc As String)
    Dim Email As Object
    Dim oMails As TB_Context.TBMails
    Dim Index As Long, Index2 As Long
    Dim sMail As String
    
    Const NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
    
    Index = ThisSends.Lookup(TaskNumber)
    Set Email = CreateObject("CDO.Message")
    
    Email.From = ThisSmtp.Sender
    
    Set oMails = ThisSends.ToMail(Index)
    For Index2 = 0 To oMails.Size - 1
        sMail = sMail & oMails.Mail(Index2) & ","
    Next Index2
    Set oMails = Nothing
    If Len(sMail) > 0 Then sMail = Left(sMail, Len(sMail) - 1)
    Email.To = sMail: sMail = ""
    
    Set oMails = ThisSends.CCMail(Index)
    For Index2 = 0 To oMails.Size - 1
        sMail = sMail & oMails.Mail(Index2) & ","
    Next Index2
    Set oMails = Nothing
    Email.Cc = sMail: sMail = ""
    
    Email.Subject = TaskDesc & " -Log Report " & Format(Now, "YYYY/MM/DD hh:mm:ss")
    Email.Htmlbody = sHml
'    Email.Textbody = sHml

    Email.AddAttachment locDir

    With Email.Configuration.Fields
        .Item(NameSpace & "sendusing") = 2
        .Item(NameSpace & "smtpserver") = ThisSmtp.Smtp
        .Item(NameSpace & "smtpserverport") = ThisSmtp.Port
        .Item(NameSpace & "smtpauthenticate") = 1
        .Item(NameSpace & "sendusername") = ThisSmtp.Username
        .Item(NameSpace & "sendpassword") = ThisSmtp.Password
        .Update
    End With

    Email.Send
    Set Email = Nothing
End Sub

Private Sub SetSend(rs As ADODB.Recordset)
    Dim Index As Long
    Dim sSQL As String
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        sSQL = "UPDATE " & CONST_TABLENAME_LOG & " SET FIsSend = 2 WHERE FID IN(0"
        
        For Index = 1 To rs.RecordCount
            sSQL = sSQL & "," & rs("FID")
            rs.MoveNext
        Next Index
        
        sSQL = sSQL & ")"
        
        TB_Runtime.ExecuteTransSQL sSQL
        sSQL = ""
    End If
End Sub

Public Sub LogEx(ByVal Info As String)
    Dim TaskNumber As String
    TaskNumber = "TASK" & ThisParameter.Value("@TASK")
    TB_Runtime.Log TaskNumber, Info
End Sub
