VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event SetBarMaxValue(ByVal Value As Long)
Public Event SetBarMinValue(ByVal Value As Long)
Public Event SetBarValue(ByVal Value As Long)
Public Event SetMsg(ByVal Msg As String)
Public Event SetBarValueWithMax()
Public Event Unload()
Public Event Active()
Public Event HideProgBar()
Public Event ShowProgBar()

Public Function SetBarMaxValue(ByVal Value As Long)
    RaiseEvent SetBarMaxValue(Value)
End Function

Public Function SetBarMinValue(ByVal Value As Long)
    RaiseEvent SetBarMinValue(Value)
End Function

Public Function SetBarValue(ByVal Value As Long)
    RaiseEvent SetBarValue(Value)
End Function
Public Function SetBarValueWithMax()
    RaiseEvent SetBarValueWithMax
End Function

Public Function SetMsg(ByVal Msg As String)
    RaiseEvent SetMsg(Msg)
End Function
Public Function SetActive()
    RaiseEvent Active
End Function
Public Function Unload()
    RaiseEvent Unload
End Function
Public Function ShowProgBar()
    RaiseEvent ShowProgBar
End Function
Public Function HideProgBar()
    RaiseEvent HideProgBar
End Function


