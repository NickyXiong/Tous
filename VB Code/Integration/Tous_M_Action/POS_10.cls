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
    
    Dim strSQL As String
    Dim StrSqlTemp As String
    
On Error GoTo ImportStockImage_Error

    InitDataEnv sDsn
    Set CNN = datasource.Connection

    For i = LBound(varDatas) To UBound(varDatas)
        If i = 0 Then
            '检查仓库
            lFStockID = GetItemID(CNN, CStr(varDatas(i)(0)), 5)
            If lFStockID = 0 Then
                 sErr = "Warehouse Master Data [" & CStr(varDatas(i)(0)) & "] is NOT created in ERP system."
                 GoTo ImportStockImage_Error
             End If
                 
            '检查是否存在，存在则将原有数据先删除
            strSQL = "select 1 from t_Tous_StockImage where FStockID=" & lFStockID & " and CONVERT(varchar(100),FDate, 23)=CONVERT(varchar(100)," & CStr(varDatas(i)(1)) & ", 23)"
            Set rs = CNN.Execute(strSQL)
            If rs.RecordCount > 0 Then
                strSQL = "delete from t_Tous_StockImage where FStockID=" & lFStockID & " and CONVERT(varchar(100),FDate, 23)=CONVERT(varchar(100)," & CStr(varDatas(i)(1)) & ", 23)"
                CNN.Execute (strSQL)
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
         
        strSQL = strSQL & vbCrLf & "insert t_Tous_StockImage (FDate,FItemID,FStockID,FSPID,FUPC,FQty)"
        strSQL = strSQL & vbCrLf & "values('" & CDate(varDatas(i)(1)) & "'," & lFItemID & "," & lFStockID & "," & lSPID & ",'" & CStr(varDatas(i)(2)) & "'," & Val(varDatas(i)(3)) & ")"
    Next

    CNN.Execute strSQL
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








