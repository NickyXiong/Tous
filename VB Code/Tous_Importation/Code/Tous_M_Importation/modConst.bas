Attribute VB_Name = "modConst"
Option Explicit

Public Const WgRk = 1   '采购入库
Public Const XcRk = 6   '虚仓入库
Public Const XsCk = 21  '销售出库
Public Const XcCk = 26  '虚仓出库
Public Const DBD = 41   '调拨
Public Const DhJh = 68 '配货单
Public Const CgD = 71   '采购订单
Public Const XCDBD = 74 '虚仓调拨
Public Const CgFpz = 75 '采购发票(专用)
Public Const CgFpp = 76 '采购发票(普通)
Public Const XsFpz = 80 '销售发票(专用)
Public Const XsD = 81   '销售订单
Public Const THTZ = 82  '退货通知单
Public Const FHTZ = 83  '发货通知单
Public Const XsFpp = 86 '销售发票(普通)


'配货调拨类型
Public Const FOFI = 1       '调出调入仓库直接取配货单调出调入仓库
Public Const SFOZI = 2      '供货方分仓调入总仓
Public Const RZOFI = 3      '接收方总仓调入分仓
Public Const SZOFI = 4      '供货方总仓调入分仓
Public Const RFOZI = 5      '接收方分仓调入总仓


Public Const C_CHECKBILL = "K3DefineBill.BillTemplateInfo"
Public Const C_BILLDATAACCESS = "BillDataAccess.GetData"
'Private Const C_BILLPACKAGE = "K3Bills.clsBillPackage"
'新加的中间层组件
Public Const C_MBILLPACKAGE = "BillDataAccess.clsBillPackage"
Public Const C_ListUpdate = "K3ListServer.clsListUpdate"
