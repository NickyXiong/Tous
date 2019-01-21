VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "POS_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************
'Function: 导入库存镜像 (POS TO ERP)
'Date:
'*******************************************************

Option Explicit

Private datasource As CDataSource
Private m_oParse As CParse
Private strLanguage As String

Private Function InitDataEnv(ByVal sToParse As String) As Boolean
    Set m_oParse = New CParse
    If m_oParse.ParseString(sToParse) Then
        Set datasource = New CDataSource
        Set datasource.ParseObject = m_oParse
    Else
        Err.Raise EBS_E_TypeMismatch, "ParseString"
    End If
    strLanguage = datasource.ParseObject.LanguageID
End Function

'---------------------------------------------------------------------------------------
' Procedure : ImportStockImage
' DateTime  :
' Author    :
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function ImportStockImage(ByVal sDsn As String, varDatas As Variant, ByRef lInterID As Long, ByRef sBillNo As String, ByRef sErr As String) As Boolean
    Dim CNN As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim rsPOInfo As ADODB.Recordset
    
    Dim i As Integer
    Dim lFStockID As Long, lFItemID As Long, lSPID As Long, lUnitID As Long
    
    Dim strSql As String
    Dim strSQLtemp As String
    
    Dim strAllSQL As New StringBuilder
    
On Error GoTo ImportStockImage_Error

    InitDataEnv sDsn
    Set CNN = datasource.Connection

    strSQLtemp = "set nocount on"
    strSQLtemp = strSQLtemp & vbCrLf & "insert t_Tous_StockImage (FDate,FItemID,FStockID,FSPID,FUPC,FQty)"
    strSQLtemp = strSQLtemp & vbCrLf & "select CONVERT(varchar(100), GETDATE(), 23),t1.FItemID,FStockID,FStockPlaceID,t2.FNumber,FQty"
    strSQLtemp = strSQLtemp & vbCrLf & "from ICInventory t1 inner join t_ICItem t2 on t1.FItemID=t2.FItemID"
    CNN.Execute (strSQLtemp)
           
    strSql = "set nocount on"
    
    For i = LBound(varDatas) To UBound(varDatas)
        If i = 0 Then
           '检查仓库
           lFStockID = GetItemID(CNN, CStr(varDatas(i)(0)), 5)
           If lFStockID = 0 Then
                sErr = "Warehouse Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                GoTo ImportStockImage_Error
            End If
                
           '检查是否存在，存在则将原有数据先删除
           strSQLtemp = "select 1 from t_Tous_POSStockImage where CONVERT(varchar(100),FDate, 23)=CONVERT(varchar(100),'" & CStr(varDatas(i)(1)) & "', 23)"
           Set rs = CNN.Execute(strSQLtemp)
           If rs.RecordCount > 0 Then
               strSQLtemp = "delete from t_Tous_POSStockImage where CONVERT(varchar(100),FDate, 23)=CONVERT(varchar(100),'" & CStr(varDatas(i)(1)) & "', 23)"
               CNN.Execute (strSQLtemp)
           End If
        End If
        
        '检查物料是否存在
        lFItemID = GetItemIDByNumber(CNN, CStr(varDatas(i)(2)), 4, lUnitID)
        If lFItemID = 0 Then
            sErr = "Material Master Data [" & CStr(varDatas(i)(2)) & "] is NOT created in ERP system."
            GoTo ImportStockImage_Error
        End If
        
        '仓位暂时设置为0
        lSPID = 0
        
        strSql = strSql & vbCrLf & "insert t_Tous_POSStockImage (FDate,FItemID,FStockID,FSPID,FUPC,FQty)"
        strSql = strSql & vbCrLf & "values('" & CDate(varDatas(i)(1)) & "'," & lFItemID & "," & lFStockID & "," & lSPID & ",'" & CStr(varDatas(i)(2)) & "'," & Val(varDatas(i)(3)) & ")"
        
        If i + 1 Mod 50 = 0 Then
            CNN.Execute strSql
            strSql = "set nocount on"
        End If
        
    Next
    
    If strSql <> "set nocount on" Then
        CNN.Execute strSql
    End If
    
    Set CNN = Nothing
    ImportStockImage = True
    GetObjectContext.SetComplete

   Exit Function

ImportStockImage_Error:
    ImportStockImage = False
    sErr = "Import Stock Image Error：" & sErr & Err.Description
    GetObjectContext.SetAbort
    Set CNN = Nothing
    

End Function








