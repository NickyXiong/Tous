VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TYPE_TB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ENUM_FTPROOT
        ROOT_K3 = 0
        ROOT_TBHQ = 2
        ROOT_3PL = 4
        ROOT_POS = 8
End Enum

Public Enum ENUM_FTPRANGE
        RANGE_LOCAL = 0
        RANGE_SINGLE = 1
        RANGE_UNION = 2
End Enum

Public Enum ENUM_RUNSTYLE
        RUNSTYLE_ACTUAL = 1
        RUNSTYLE_FIXED = 2
End Enum

Public Type TYPE_EMAILSMTP
        Smtp As String
        Username As String
        Password As String
        Port As Integer
        Sender As String
End Type

Public Type TYPE_K3SERVER
        DBServer As String
        DBName As String
        DBUsername As String
        DBPassword As String
        K3Username As String
End Type

Public Type TYPE_PARAMETER
        ID As Long
        Name As String
        Key As String
        Value As String
        Description As String
End Type

Public Type TYPE_REMOTE
        RootID As ENUM_FTPROOT
        Name As String
        RangeID As ENUM_FTPRANGE
        DownRoot As String
        UpRoot As String
        CacheRoot As String
        BackupRoot As String
        Server As String
        Username As String
        Password As String
        Description As String
End Type

Public Type TYPE_METADATA
        ID As String
        FromID As ENUM_FTPROOT
        ToID As ENUM_FTPROOT
        Filename As String
        DtFormat As String
        SnFormat As String
        Postfix As String
        Directory As String
End Type

Public Type TYPE_TASK
        Number As String
        IsSys As Boolean
        IsUse As Boolean
        ExeName As String
        Parameter As String
        RunStyle As ENUM_RUNSTYLE
        Interval As String
        StartTime As String
        StopTime As String
        Description As String
End Type

Public Type TYPE_ACTION
        TaskNumber As String
        ActionNumber As String
        MetadataNumber As String
        DllName As String
End Type

Public Type TYPE_ACTION_DATA
        TaskNumber As String
        ActionNumber As String
        MetadataNumber As String
        FromID As ENUM_FTPROOT
        ToID As ENUM_FTPROOT
        Filename As String
        CacheRoot As String
        BackupRoot As String
        Root As String
End Type
