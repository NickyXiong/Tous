VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HSXMPlug_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'K3物料基础资料插件

Private WithEvents bObj As K3BaseDataEvents.ItemEvents
Attribute bObj.VB_VarHelpID = -1
Private baseObj As K3BaseDataEvents.ItemEditInterface
Attribute baseObj.VB_VarHelpID = -1

Private lProductName As Long
Private lProductNumber As Long

Private lCustomer As Long

Public Sub Show(ByVal o As K3BaseDataEvents.ItemEvents)
'    MMTS.CheckMts 1
    Set bObj = o
    Set baseObj = o.EditInterface
    
    
    Select Case baseObj.ItemClassId
        Case 3009
            lProductName = baseObj.GetRowNumber("F_112")
            lProductNumber = baseObj.GetRowNumber("F_114")
            lCustomer = baseObj.GetRowNumber("F_106")
            baseObj.SetRowLock lProductNumber
        Case 3001
            lCustomer = baseObj.GetRowNumber("F_101")
        Case 3008
            lCustomer = baseObj.GetRowNumber("F_102")
            
    End Select

End Sub


Private Sub bObj_AfterGridLoad()
    Dim strSQL As String
    Dim rs As adodb.Recordset

    Dim strParentNumber As String

    If baseObj.ItemIsDetail = True Then
        If Not baseObj.ParentItemObject Is Nothing Then
            strParentNumber = baseObj.ParentItemObject.Number
            strSQL = "select FItemID,FName from t_Organization where FNumber='" & strParentNumber & "'"
            Set rs = modPub.ExecuteSql(strSQL)
            If rs.RecordCount > 0 Then
                baseObj.SetRowValue lCustomer, rs.Fields("FName").Value, rs.Fields("FItemID").Value
    '            baseObj.SetRowValue lCustomer, strParentNumber
                baseObj.SetRowLock lCustomer
            End If
        End If
    End If
End Sub



Private Sub bObj_BeforeSave(Cancel As Boolean)
    Dim strNewNumber As String
    Dim strParentNumber As String
    
    Dim lCustomerID As Long
    Dim strSQL As String
    Dim rs As adodb.Recordset
    

    strNewNumber = CStr(baseObj.ReadValue(baseObj.GetRowNumber("FNumber")))
    
    If baseObj.ItemIsDetail = True Then
        If Not baseObj.ParentItemObject Is Nothing Then
            strParentNumber = baseObj.ParentItemObject.Number
            
            If InStr(1, strNewNumber, strParentNumber & ".") <> 1 Then
                baseObj.SetRowValue baseObj.GetRowNumber("FNumber"), strParentNumber & "." & strNewNumber
            End If
        Else
            lCustomerID = baseObj.ReadID(lCustomer)
            
            If lCustomerID <> 0 Then
                strSQL = "select FNumber from t_Organization where fitemid=" & lCustomerID
                Set rs = modPub.ExecuteSql(strSQL)
                
                If rs.RecordCount > 0 Then
                    baseObj.SetRowValue baseObj.GetRowNumber("FNumber"), rs.Fields("FNumber").Value & "." & strNewNumber
                End If
            End If
        End If
    End If
End Sub

Private Sub bObj_GridChange(ByVal lRow As Long)
    Dim lItemID As Long
    Dim strSQL As String
    Dim rs As adodb.Recordset
    
    If baseObj.ItemIsDetail = True Then
        If baseObj.ItemClassId = 3009 Then
            If lRow = lProductName Then
                lItemID = baseObj.ReadID(lRow)
                
                strSQL = "select FNumber from t_icitem where fitemid=" & lItemID
                Set rs = modPub.ExecuteSql(strSQL)
                
                baseObj.SetRowValue lProductNumber, rs.Fields("FNumber").Value
            End If
        End If
    End If
End Sub


