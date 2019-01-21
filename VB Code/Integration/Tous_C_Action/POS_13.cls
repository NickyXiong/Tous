VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POS_13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : POS_13
' Purpose   : 导入POS的维修单数据生成收款单
'---------------------------------------------------------------------------------------

Option Explicit

Implements TB_Runtime.ITask

Private m_sDsn As String
Private m_sPath As String
Private m_sFileName As String
Private m_sFullPath As String
Private m_sActionNumer As String
Private m_sTaskNumber As String
Private m_data As TB_Context.TYPE_ACTION_DATA
Private Sub ITask_Dispose(param As TB_Context.TBParameters)

End Sub

Private Function ITask_Init(k3svr As TB_Context.TYPE_K3SERVER, data As TB_Context.TYPE_ACTION_DATA, param As TB_Context.TBParameters) As Boolean
    
   On Error GoTo ITask_Init_Error

    m_sPath = data.CacheRoot
    m_sFileName = data.FileName
    m_sFullPath = m_sPath & m_sFileName
    m_sActionNumer = data.ActionNumber
    m_sTaskNumber = data.TaskNumber
    
    m_data = data
    m_sDsn = "ConnectString={" & K3ConnectionString(k3svr) & "};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=;AcctType=gy;Setuptype=Industry;Language=chs;IP=127.0.0.1;K3Version=KUE;MachineName=127.0.0.1;UUID=BAD160C5-B4DB-4965-A8A7-0D373C885AE6"
    ITask_Init = True
    
   Exit Function

ITask_Init_Error:
    ITask_Init = False
    
    TB_Runtime.Log data.TaskNumber, "Tous_C_Action.POS_13.ITask_Init:" & Err.Description
    
End Function



Private Function ITask_Run(param As TB_Context.TBParameters) As Boolean
    Dim blnSyn As Boolean
On Error GoTo ITask_Run_Error

    blnSyn = SynBill1000005
    ITask_Run = blnSyn
   
    Exit Function

ITask_Run_Error:

    TB_Runtime.Log m_data.TaskNumber, "Tous_C_Action.POS_13.ITask_Run:" & Err.Description
End Function


Public Function SynBill1000005() As Boolean


    Dim vData As Noteset
    Dim i As Long
    Dim obj As Object
    Dim blnSave As Boolean
    Dim vBill As Variant
    Dim sError As String
    Dim lInterID As Long
    Dim sBillNo As String
    Dim v(2) As Variant
    
   On Error GoTo SynBill1000005_Error

    Set vData = New Noteset
    v(0) = 1
    v(1) = 2
    v(2) = 3
'    v(3) = 10
    
    vData.LoadData m_sFullPath, v
    
    While vData.IndexNext
        
        vBill = vData.GetDataByIndex
        If UboundEx(vBill) <> -1 Then
            Set obj = CreateObject("Tous_M_Action.POS_13")
            lInterID = 0
            sBillNo = ""
            sError = ""
            blnSave = obj.SynBill1000005(m_sDsn, vBill, sBillNo, sError)
            
            If blnSave = True Then
                If sError = "" Then
                    TB_Runtime.K3Log m_data, "Import successfully:POS No[" & vBill(0)(1) & "]", False
                Else
                    TB_Runtime.K3Log m_data, "Import failed:" & Left(sError, 980), True
                    SynBill1000005 = False
                End If
            Else
                TB_Runtime.K3Log m_data, "Import failed:" & Left(sError, 980), True
                SynBill1000005 = False
            End If

            Set obj = Nothing
        End If
    Wend
    
    Erase vBill
    Set vData = Nothing
    SynBill1000005 = True
    Exit Function

SynBill1000005_Error:
    
    Erase vBill
    Set vData = Nothing
    SynBill1000005 = False
    TB_Runtime.Log m_sTaskNumber, "Tous_C_Action.POS_13.SynBill1000005:" & Err.Description
End Function






