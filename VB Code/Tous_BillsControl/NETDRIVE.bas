Attribute VB_Name = "NETDRIVE"
Option Explicit
'*********************网络驱动器定义开始               ***************************************'
  
'添加到网络驱动器的连接
Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias _
          "WNetAddConnection2A" (lpNetResource As NETRESOURCE, _
          ByVal lpPassword As String, ByVal lpUserName As String, _
          ByVal dwFlags As Long) As Long
            
'取消到网络驱动器的连接
Public Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias _
          "WNetCancelConnection2A" (ByVal lpName As String, _
          ByVal dwFlags As Long, ByVal fForce As Long) As Long
  
Public Const NO_ERROR = 0
Public Const CONNECT_UPDATE_PROFILE = &H1
  
'网络驱动器参数
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_REMEMBERED = &H3
Public Const RESOURCE_GLOBALNET = &H2
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
'错误常量
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_ALREADY_ASSIGNED = 85&
Public Const ERROR_BAD_DEV_TYPE = 66&
Public Const ERROR_BAD_DEVICE = 1200&
Public Const ERROR_BAD_NET_NAME = 67&
Public Const ERROR_BAD_PROFILE = 1206&
Public Const ERROR_BAD_PROVIDER = 1204&
Public Const ERROR_BUSY = 170&
Public Const ERROR_CANCELLED = 1223&
Public Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Public Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Public Const ERROR_EXTENDED_ERROR = 1208&
Public Const ERROR_INVALID_PASSWORD = 86&
Public Const ERROR_NO_NET_OR_BAD_PATH = 1203&
  
  
'网络驱动器映射
Public Type NETRESOURCE
        dwScope   As Long
        dwType   As Long
        dwDisplayType   As Long
        dwUsage   As Long
        lpLocalName   As String
        lpRemoteName   As String
        lpComment   As String
        lpProvider   As String
End Type
'定义一个全局的本地网络驱动器变量（当网络驱动器连接的时候自动更新，同时要检测是否关闭原来的网络驱动器）
Public LocalNetDrive     As String
'*********************网络驱动器定义结束               ***************************************'


'连接到网络驱动器
Public Function NetDriveConnect(ByVal RemotePath As String, ByVal Localpath As String, ByVal lpUserName As String, ByVal lpPassword As String) As Boolean
          
        NetDriveConnect = False
        Dim NetR     As NETRESOURCE
        Dim ErrInfo     As Long
          
        On Error GoTo Error_NetDriveConnect
        Dim f As New FileSystemObject
        If f.DriveExists(Localpath) = True Then           '如果该磁盘已经存在，就不再重新建立连接
                NetDriveConnect = True
                LocalNetDrive = Localpath
        Else
                NetR.dwScope = RESOURCE_GLOBALNET
                NetR.dwType = RESOURCETYPE_DISK
                NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
                NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
                NetR.lpLocalName = Localpath
                LocalNetDrive = Localpath
                NetR.lpRemoteName = RemotePath
                ErrInfo = WNetAddConnection2(NetR, lpPassword, lpUserName, CONNECT_UPDATE_PROFILE)             '用户名和密码
                If ErrInfo = NO_ERROR Then NetDriveConnect = True
        End If
        Exit Function
          
Error_NetDriveConnect:
        NetDriveConnect = False
End Function

'断开网络驱动器
Public Function NetDriveDisconnect(ByVal LocalNetDrive As String) As Boolean
  
        NetDriveDisconnect = False
        Dim ErrInfo     As Long
          
        On Error GoTo Error_NetDriveDisconnect
          
        ErrInfo = WNetCancelConnection2(LocalNetDrive, CONNECT_UPDATE_PROFILE, True)
        If ErrInfo = NO_ERROR Then NetDriveDisconnect = True
          
        Exit Function
          
Error_NetDriveDisconnect:
    NetDriveDisconnect = False
      
End Function

