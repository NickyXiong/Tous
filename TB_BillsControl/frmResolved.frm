VERSION 5.00
Begin VB.Form frmResolved 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resolved Remark"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   Icon            =   "frmResolved.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtRemark 
      Height          =   375
      Left            =   120
      MaxLength       =   225
      TabIndex        =   1
      Top             =   720
      Width           =   5895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "(Max length: 225 Characters)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Remark"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmResolved"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lngInterID As String
Public m_strRemark As String
Public m_strStatus As String
Public m_bUpdated As Boolean

Public m_bNewBill As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim respon
    Dim strRemark As String
    Dim strSQL As String
    
    If m_bNewBill = True Then
        If m_strStatus <> "Y" Then
            respon = MsgBox("Do you want to resolve this transaction bill?", vbYesNo, "Kingdee Prompt")
            If respon = vbYes Then
                strRemark = "Resolved by " & lblUser.Caption & " (" & Format(Now(), "yyyy/mm/dd hh:mm") & ")."
                strRemark = strRemark & txtRemark.Text & vbCrLf & m_strRemark
                
                strRemark = ChangeStr(strRemark)
                
                strSQL = "update t_TB_StockDeliveryNotice set FResolvedStatus='Y',FResolvedRemark='" & strRemark & "' "
                strSQL = strSQL & "where FID =" & lngInterID
                modPub.ExecuteSql (strSQL)
                MsgBox "Resolve transaction bill successfully", vbInformation, "Kingdee Prompt"
                m_strRemark = strRemark
                m_bUpdated = True
            End If
        Else
            respon = MsgBox("Do you want to Un-resolve this transaction bill?", vbYesNo, "Kingdee Prompt")
            If respon = vbYes Then
                strRemark = "Un-resolved by " & lblUser.Caption & " (" & Format(Now(), "yyyy/mm/dd hh:mm") & ")."
                strRemark = strRemark & txtRemark.Text & vbCrLf & m_strRemark
                
                strRemark = ChangeStr(strRemark)
                
                strSQL = "update t_TB_StockDeliveryNotice set FResolvedStatus='N',FResolvedRemark='" & strRemark & "' "
                strSQL = strSQL & "where FID =" & lngInterID
                modPub.ExecuteSql (strSQL)
                MsgBox "Un-resolve transaction bill successfully", vbInformation, "Kingdee Prompt"
                m_strRemark = strRemark
                m_bUpdated = True
            End If
        End If
    Else
        If m_strStatus <> "Y" Then
            respon = MsgBox("Do you want to resolve this transaction bill?", vbYesNo, "Kingdee Prompt")
            If respon = vbYes Then
                strRemark = "Resolved by " & lblUser.Caption & " (" & Format(Now(), "yyyy/mm/dd hh:mm") & ")."
                strRemark = strRemark & txtRemark.Text & vbCrLf & m_strRemark
                
                strRemark = ChangeStr(strRemark)
                
                strSQL = "update icstockbill set FResolvedStatus='Y',FResolvedRemark='" & strRemark & "' "
                strSQL = strSQL & "where FTranType=41 and FInterID =" & lngInterID
                modPub.ExecuteSql (strSQL)
                MsgBox "Resolve transaction bill successfully", vbInformation, "Kingdee Prompt"
                m_strRemark = strRemark
                m_bUpdated = True
            End If
        Else
            respon = MsgBox("Do you want to Un-resolve this transaction bill?", vbYesNo, "Kingdee Prompt")
            If respon = vbYes Then
                strRemark = "Un-resolved by " & lblUser.Caption & " (" & Format(Now(), "yyyy/mm/dd hh:mm") & ")."
                strRemark = strRemark & txtRemark.Text & vbCrLf & m_strRemark
                
                strRemark = ChangeStr(strRemark)
                
                strSQL = "update icstockbill set FResolvedStatus='N',FResolvedRemark='" & strRemark & "' "
                strSQL = strSQL & "where FTranType=41 and FInterID =" & lngInterID
                modPub.ExecuteSql (strSQL)
                MsgBox "Un-resolve transaction bill successfully", vbInformation, "Kingdee Prompt"
                m_strRemark = strRemark
                m_bUpdated = True
            End If
        End If
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    lblUser.Caption = MMTS.UserName
    m_bUpdated = False
    If m_strStatus <> "Y" Then
        Me.Caption = "Resolved Remark"
    Else
        Me.Caption = "Un-resolved Remark"
    End If
End Sub
