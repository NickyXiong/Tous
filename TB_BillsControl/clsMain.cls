VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CallList(ByVal ListID As Long _
    , Optional ByVal bShowFilter As Boolean = False _
    , Optional ByVal sFilter As String = "")
    Dim obj As clsOldBill
    Set obj = New clsOldBill
    obj.OpenList ListID, bShowFilter, sFilter
    Set obj = Nothing
End Function

Public Function CallNewList(ByVal lClassTypeID As Long _
    , Optional ByVal bShowFilter As Boolean = True _
    , Optional ByVal sHeadSort As String = "" _
    , Optional ByVal sEntrySort As String = "" _
    , Optional ByVal sSelPage As String = "" _
    , Optional ByVal sFilter As String = "")

    Dim obj As clsNewBill
    Set obj = New clsNewBill
    obj.CallBosList lClassTypeID, bShowFilter, sHeadSort, sEntrySort, sSelPage, sFilter
    Set obj = Nothing
End Function

Public Function CallRpt()

End Function

Public Function GetList(ByVal ListID As Long _
    , Optional ByVal bShowFilter As Boolean = False _
    , Optional ByVal sFilter As String = "" _
    ) As KFO.Vector
    Dim obj As clsOldBill
    Set obj = New clsOldBill
    Set GetList = obj.LookUpList(ListID, bShowFilter, sFilter)
    Set obj = Nothing
    
End Function

Public Function GetNewList(ByVal lClassTypeID As Long _
    , Optional ByVal bShowFilter As Boolean = True _
    , Optional ByVal sHeadSort As String = "" _
    , Optional ByVal sEntrySort As String = "" _
    , Optional ByVal sSelPage As String = "" _
    , Optional ByVal sFilter As String = "" _
    , Optional ByVal bSelWhole As Boolean = False) As KFO.Vector
    Dim obj As clsNewBill
    Set obj = New clsNewBill
    Set GetNewList = obj.LookUpBosList(lClassTypeID, bShowFilter, sHeadSort, sEntrySort, sSelPage, sFilter, bSelWhole)
    Set obj = Nothing
End Function

Public Function CallNewBill(ByVal lClassTypeID As Long)
    Dim obj As clsNewBill
    Set obj = New clsNewBill
    obj.CallBosBill lClassTypeID
    Set obj = Nothing
End Function

Private Sub Class_Initialize()
'    MMTS.CheckMts True
End Sub
