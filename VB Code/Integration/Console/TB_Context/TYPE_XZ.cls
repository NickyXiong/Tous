VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TYPE_XZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ENUM_XZ_PROCESSSTATE
        STATE_NOMAL = 0
        STATE_RUN = 1
        STATE_STOP = 2
End Enum

Public Type TYPE_XZ_PROCESSATTRIBUTE
        Key As String
        Directory As String
        Command As String
        Mask As String
        PHD As Long
        THD As Long
        PID As Long
        TID As Long
End Type

Public Type TYPE_XZ_ERROR
        ErrNumber As String
        ErrSource As String
        ErrDescription As String
End Type

Public Type TYPE_XZ_SHARE
        Identity As String
        Resource As String
End Type
