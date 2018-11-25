VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BosDataHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oData As KFO.Dictionary
Private oTableInfo As KFO.Dictionary
Private billClassId As Long


Public Property Set PackedData(data As KFO.Dictionary)
    Set oData = data
    
    Dim ibos As Object
    Set ibos = CreateObject("K3ClassTpl.DataSrv")
    ibos.propstring = MMTS.PropsString
    ibos.ClassTypeID = data("FClassTypeID")
    billClassId = data("FClassTypeID")
    
    Set oTableInfo = ibos.TableInfo
    
End Property

Public Property Get PackedData() As KFO.Dictionary
    Set PackedData = oData
End Property



Public Property Get ClassTypeID() As Long
    ClassTypeID = billClassId
    
End Property


Public Property Get TableInfo() As KFO.Dictionary
    Set TableInfo = oTableInfo
End Property

Public Sub NewEmptyBill(dsn As String, billClasstypeId As Long)
    Dim ibos As Object
    Set ibos = CreateObject("K3ClassTpl.DataSrv")
    ibos.propstring = dsn
    ibos.ClassTypeID = billClasstypeId
    billClassId = billClasstypeId
    
    Set oTableInfo = ibos.TableInfo
    Set oData = ibos.GetEmptyBill
    
    billClassId = billClasstypeId
End Sub

Public Sub SetHValue(key As String, value As Variant)
    oData("Page1")(oTableInfo("map")(key))("FFLD") = value
End Sub


Public Sub SetTValue(page As String, row As Long, key As String, value As Variant)
    oData(page)(row)(oTableInfo("map")(key))("FFLD") = value
End Sub

Public Sub AddNewEntryRow(page As String)
    oData(page).Add oData(page)(1).CloneAll
End Sub


Public Function GetHValue(key As String) As Variant
    GetHValue = oData("Page1")(oTableInfo("map")(key))("FFLD")
End Function

Public Function GetTValue(page As String, row As Long, key As String) As Variant
    GetTValue = oData(page)(row)(oTableInfo("map")(key))("FFLD")
End Function

Public Function GetRecordCount(page As String) As Long
    GetRecordCount = oData(page).Size
End Function

Public Function SaveBill() As Long
    Dim ibos As Object
    Set ibos = CreateObject("K3ClassTpl.DataSrv")
    ibos.propstring = MMTS.PropsString
    ibos.ClassTypeID = ClassTypeID
        
    Dim oSave As Object
    Set oSave = CreateObject("K3MClass.BillUpdate")
    
    SaveBill = oSave.Save(MMTS.PropsString, ibos.classType, ibos.classTypeEntry, ibos.TableInfo, oData, ibos.dctLink)
End Function
