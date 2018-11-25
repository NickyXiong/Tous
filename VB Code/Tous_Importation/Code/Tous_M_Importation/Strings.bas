Attribute VB_Name = "Strings"
' Kingdee Enterprise Business Objects
' Copyright (C) 1995-1998 Kingdee Corporation
' All rights reserved

Option Explicit
'------------------------------------------------------------------------
' System Build-in Classes
'------------------------------------------------------------------------
Public datasource As CDataSource
'Currency
Public Const Colliated_Currency = 0
Public Const Base_Currency = 1


'------------------------------------------------------------------------
' Table names

Public Const t_Account = "t_Account"
'Public Const t_AccountItem = "t_AccountItem"
Public Const t_AcctGroup = "t_AcctGroup"
Public Const t_Balance = "t_Balance"
Public Const t_Budget = "t_Budget"
Public Const t_Currency = "t_Currency"
Public Const t_Department = "t_Department"
Public Const t_Dict = "t_Dict"
Public Const t_Item = "t_Item"
Public Const t_ItemClass = "t_ItemClass"
Public Const t_MeasureUnit = "t_MeasureUnit"
Public Const t_Organization = "t_Organization"
Public Const t_Personnel = "t_Personnel"
Public Const t_ProfitAndLoss = "t_ProfitAndLoss"
Public Const t_QuantityBalance = "t_QuantityBalance"
Public Const t_RateAdjust = "t_RateAdjust"
Public Const t_UnitGroup = "t_UnitGroup"
Public Const t_Voucher = "t_Voucher"
Public Const t_VoucherDetail = "t_VoucherDetail"
Public Const t_VoucherEntry = "t_VoucherEntry"
Public Const t_VoucherGroup = "t_VoucherGroup"

Public Const t_ItemBalance = "t_ItemBalance"
Public Const t_ItemBalanceDetail = "t_ItemBalanceDetail"

Public Const t_AutoTransfer = "t_AutoTransfer"
Public Const t_AutoTransferEntry = "t_AutoTransferEntry"
'------------------------------------------------------------------------
' Column names

Public Const FAccount = "FAccount"
Public Const FAccountItem = "FAccountItem"
Public Const FAccountID = "FAccountID"
Public Const FAccountID2 = "FAccountID2"
Public Const FActualType = "FActualType"
Public Const FActualSize = "FActualSize"
Public Const FAddress = "FAddress"
Public Const FAdjustRate = "FAdjustRate"
Public Const FAmount = "FAmount"
Public Const FAmountFor = "FAmountFor"
Public Const FAttachments = "FAttachments"
Public Const FAuxClass = "FAuxClass"
Public Const FBalChecked = "FBalChecked"
Public Const FBank = "FBank"
Public Const FBeginBalance = "FBeginBalance"
Public Const FBeginBalanceFor = "FBeginBalanceFor"
Public Const FBeginQty = "FBeginQty"
Public Const FBehavior = "FBehavior"
Public Const FBirthday = "FBirthday"
Public Const FBudget = "FBudget"
Public Const FBudgetFor = "FBudgetFor"
Public Const FMinBudget = "FMinBudget"
Public Const FMinBudgetFor = "FMinBudgetFor"
Public Const FCashierID = "FCashierID"
Public Const FChecked = "FChecked"
Public Const FCheckerID = "FCheckerID"
Public Const FCity = "FCity"
Public Const FCoefficient = "FCoefficient"
Public Const FContact = "FContact"
Public Const FCountry = "FCountry"
Public Const FCredit = "FCredit"
Public Const FCreditFor = "FCreditFor"
Public Const FCreditLimit = "FCreditLimit"
Public Const FCreditQty = "FCreditQty"
Public Const FCreditTotal = "FCreditTotal"
Public Const FCurrencyID = "FCurrencyID"
Public Const FDataType = "FDataType"
Public Const FDate = "FDate"
Public Const FDC = "FDC"
Public Const FDetailID = "FDetailID"
Public Const FDebit = "FDebit"
Public Const FDebitFor = "FDebitFor"
Public Const FDebitQty = "FDebitQty"
Public Const FDebitTotal = "FDebitTotal"
'Public Const FDefault = "FDefault"
Public Const FDefaultUnit = "FDefaultUnit"
Public Const FDefaultUnitID = "FDefaultUnitID"
Public Const FDepartmental = "FDepartmental"
Public Const FDepartmentID = "FDepartmentID"
Public Const FDetail = "FDetail"
Public Const FDetailCount = "FDetailCount"
Public Const FDictItemID = "FDictItemID"
Public Const FDifference = "FDifference"
Public Const FDuty = "FDuty"
Public Const FEarnAccountID = "FEarnAccountID"
Public Const FEmail = "FEmail"
Public Const FEndBalance = "FEndBalance"
Public Const FEndBalanceFor = "FEndBalanceFor"
Public Const FEndBlance = "FEndBlance"
Public Const FEndBlanceFor = "FEndBlanceFor"
Public Const FEndQty = "FEndQty"
Public Const FEntryCount = "FEntryCount"
Public Const FEntryID = "FEntryID"
Public Const FExchangeRate = "FExchangeRate"
Public Const FExplanation = "FExplanation"
Public Const FGender = "FGender"
Public Const FGroup = "FGroup"
Public Const FGroupID = "FGroupID"
Public Const FHandler = "FHandler"
Public Const FHelperCode = "FHelperCode"
Public Const FHireDate = "FHireDate"
Public Const FHomePage = "FHomePage"
Public Const FID = "FID"
Public Const FImmPosted = "FImmPosted"
Public Const FIndex = "FIndex"
Public Const FInternalInd = "FInternalInd"
Public Const FIsBank = "FIsBank"
Public Const FIsCash = "FIsCash"
Public Const FIsFolder = "FIsFolder"
Public Const FItemClassID = "FItemClassID"

Global Const FItemClassName = "FName"
Global Const FBalID = "FBalID"

Public Const FItemID = "FItemID"
Public Const FJournal = "FJournal"
Public Const FLeaveDate = "FLeaveDate"
Public Const FLevel = "FLevel"
Public Const FMeasureUnitID = "FMeasureUnitID"
Public Const FName = "FName"
Public Const FNote = "FNote"
Public Const FNumber = "FNumber"
Public Const FObjectName = "FObjectName"
Public Const FOperator = "FOperator"
Public Const FOrganizational = "FOrganizational"
Public Const FOrganizationID = "FOrganizationID"
Public Const FOwnerGroupID = "FOwnerGroupID"
Public Const FParent = "FParent"
Public Const FParentID = "FParentID"
Public Const FParameter = "FParameter"
Public Const FPeriod = "FPeriod"
Public Const FPeriodRange = "FPeriodRange"
Public Const FPersonnel = "FPersonnel"
Public Const FPersonnelID = "FPersonnelID"
Public Const FPhone = "FPhone"
Public Const FPhone1 = "FPhone1"
Public Const FPhone2 = "FPhone2"
Public Const FPhone3 = "FPhone3"
Public Const FPostalCode = "FPostalCode"
Public Const FPosted = "FPosted"
Public Const FPosterID = "FPosterID"
Public Const FPrecision = "FPrecision"
Public Const FPreparerID = "FPreparerID"
Public Const FPropID = "FPropID"
Public Const FProportion = "FProportion"
Public Const FProvince = "FProvince"
Public Const FQuantities = "FQuantities"
Public Const FQuantity = "FQuantity"
Public Const FReference = "FReference"
Public Const FRootID = "FRootID"
Public Const FSerialNum = "FSerialNum"
Public Const FScale = "FScale"
Public Const FSettleTypeID = "FSettleTypeID"
Public Const FSettleNo = "FSettleNo"
Public Const FSQLColumnName = "FSQLColumnName"
Public Const FSummarized = "FSummarized"
Public Const FTaxID = "FTaxID"
Public Const FText = "FText"
Public Const FTransNo = "FTransNo"
Public Const FTransferID = "FTransferID"
Public Const FTransDate = "FTransDate"
Public Const FTranType = "FTranType"
Public Const FType = "FType"
Public Const FUnitGroupID = "FUnitGroupID"
Public Const FDeleted = "FDeleted"
Public Const FUnitPrice = "FUnitPrice"
Public Const FUnused = "FUnnsed"
Public Const FValue = "FValue"
Public Const FVoucherID = "FVoucherID"
Public Const FYtdAmount = "FYtdAmount"
Public Const FYtdAmountFor = "FYtdAmountFor"
Public Const FYtdCredit = "FYtdCredit"
Public Const FYtdCreditFor = "FYtdCreditFor"
Public Const FYtdCreditQty = "FYtdCreditQty"
Public Const FYtdDebit = "FYtdDebit"
Public Const FYtdDebitFor = "FYtdDebitFor"
Public Const FYtdDebitQty = "FYtdDebitQty"
Public Const FYear = "FYear"
Public Const FImport = "FImport"

Public Const Microsoft_SQL_Server = "Microsoft SQL Server"

'Voucher Consts
Public Const VoucherEntries = "_Entries"
Public Const VoucherEntryDetails = "_Details"

'Account Consts
Public Const AccountItems = "_Items"
Public Const ItemAccounts = "_Accounts"

'MeasureUnit Consts
Public Const MeasureUnitList = "_MeasureUnits"

'Dict Consts
Public Const DefaultPathSeparator = "\"

'Item Custom Properties Consts
Public Const ItemProperties = "_PropDesc"

'Fixed Asset System Consts
Public Const FABalances = "_Balances"
Public Const FAExtras = "_Extras"
Public Const FAAlterItems = "_AlterItems"
Public Const FAExpenseItems = "_ExpenseItems"
Public Const FACardBookValues = "_CardBookValues"
Public Const FACardValues = "_CardValues"
Public Const FACardDepartments = "_CardDepartments"
Public Const FACardExpenses = "_CardExpenses"
Public Const FACardAssetItems = "_CardAssetItems"
Public Const FACardDepreciationItems = "_CardDepreciationItems"
Public Const FACardProperties = "_CardPropDescs"
Public Const FACardPropValues = "_CardPropValues"

Public Const FAVoucherBookValues = "_VoucherBookValues"
Public Const FAVoucherValues = "_VoucherValues"
Public Const FAVoucherDepartments = "_VoucherDepartments"
Public Const FAVoucherExpenses = "_VoucherExpenses"
Public Const FAVoucherAssetItems = "_VoucherAssetItems"
Public Const FAVoucherDepreciationItems = "_VoucherDepreciationItems"

'SystemProfiles
Public Const BaseProfile = "Base"
Public Const GeneralProfile = "General"
Public Const GLProfile = "GL"
Public Const FAProfile = "FA"

'Base System Profiles
Public Const BaseAutoNumbers = "AutoNumbers"
    
'General System Profiles
Public Const GenCompanyName = "CompanyName"
Public Const GenCompanyPhone = "CompanyPhone"
Public Const GenCompanyAddress = "CompanyAddress"

'Accounting System Profiles (GLProfile)

Public Const GLVersion = "Version"
Public Const GLProgramName = "ProgramName"
Public Const GLAccountLength = "AccountLength"
Public Const GLCurrentPeriod = "CurrentPeriod"
Public Const GLCurrentYear = "CurrentYear"
Public Const GLMaxAccountLevel = "MaxAccountLevel"
Public Const GLPeriodCount = "PeriodCount"
Public Const GLPeriodDates = "PeriodDates"
Public Const GLYearDifference = "YearDifference"
Public Const GLClosed = "Closed"
Public Const GLStartYear = "StartYear"
Public Const GLStartPeriod = "StartPeriod"
Public Const GLEarnAccount = "EarnAccount"
Public Const GLPeriodByMonth = "PeriodByMonth"
Public Const GLCheckBeforePost = "CheckBeforePost"
Public Const GLRateAdjustAccount = "RateAdjustAccount"
Public Const GLRateAdjustVoucherGroup = "RateAdjustVoucherGroup"
Public Const GLRateAdjustVoucherExplanation = "RateAdjustVoucherExplanation"
Public Const GLTransPLVoucherGroup = "TransPLVoucherGroup"
Public Const GLTransPLVoucherExplanation = "TransPLVoucherExplanation"
Public Const GLInitClosed = "InitClosed"
Public Const GLEndBalDCFormat = "EndBalDCFormat"         'XYF 99-11-08 帐簿余额方向
'Fixed Asset System Profiles (FAProfile)
Public Const FAVersion = "FAVersion"
Public Const FAProgramname = "FAProgramName"
Public Const FAMaxFullNumberLength = 80
Public Const FAMaxCardFullNumberLength = 90

'InternalInd Vouchers
Public Const RateAdjustVoucher = "RateAdjust"
Public Const TransferPLVoucher = "TransferPL"

Public Const FForCurrencyColumn = "FForCurrencyColumn"
Public Const FExchangeRateColumn = "FExchangeRateColumn"
Public Const FBaseCurrencyColumn = "FBaseCurrencyColumn"
Public Const FColumnID = "FColumnID"
Public Const FColumnName = "FColumnName"
Public Const FSerialNo = "FSerialNo"
Public Const FSubColumnID = "FSubColumnID"
Public Const FSubColumnName = "FSubColumnName"
Public Const MultiColumns = "_Columns"
Public Const MultiSubColumns = "_SubColumns"
Public Const t_MultiColumnLedger = "t_MultiColumnLedger"
Public Const t_MultiColumn = "t_MultiColumn"
Public Const t_MultiSubColumn = "t_MultiSubColumn"
Public Const FCurrencyName = "FCurrencyName"
Public Const FAccountNumber = "FAccountNumber"
Public Const FLedgerID = "FLedgerID"
Public Const FLedgerName = "FLedgerName"

' CONST for Contact

' Tables
Public Const t_Contact = "t_Contact"
Public Const t_ContactEntry = "t_ContactEntry"
Public Const t_ContactVoucher = "t_ContactVoucher"

' Fields
Public Const FContactID = "FContactID"
Public Const FContactNumber = "FNumber"
Public Const FClosed = "FClosed"
Public Const FBalance = "FBalance"

Public Const ContactEntries = "_Entries"

Public Const AutoTransferEntries = "_Entries"

Public Enum VoucherEntryDC
    DC_Debit = 1
    DC_Credit = 0
End Enum

Public Enum CustomPropertySortEnum
    SortByPropID = 0
    SortByName = 1
    SortByBehavior = 2
End Enum

Public Enum AccountDCConsts
    ebAccountDC_Debit = 1
    ebAccountDC_Credit = -1
    ebAccountDC_Balance = 0
End Enum

Public Enum ItemClassBuildinConsts
    Itemclass_Organization = 1
    ItemClass_Department = 2
    ItemClass_Personnel = 3
    ItemClass_Contact = 7
    'ItemClass_SettleType = 8
End Enum
Public Enum VoucherEditMode
    vchNew = 0
    vchView = 1
    vchModify = 2
    vchCheck = 3
End Enum

Public Const vbKeyLookup = vbKeyF7

'-----------------------------------------------------------------------------------
'Settle Center System

'RateType
Public Const RTYear = 0
'Public Const RTQuarter = 2
Public Const RTMonth = 1
Public Const RTDay = 2


'------------------------------------------------------------------------
' Table names

Public Const t_Acnt = "t_Acnt"

'------------------------------------------------------------------------
' Column names
Public Const FAcctint = "FAcctint"
Public Const FintRate = "FintRate"
Public Const FAcctID = "FAcctID"
Public Const FAcctNo = "FAcctNo"
Public Const FAcnt = "FAcnt"
Public Const FAcntID = "FAcntID"
Public Const FAcntNo = "FAcntNo"
Public Const FAcntName = "FAcntName"
Public Const FAllowOD = "FAllowOD"
Public Const FARIntAcct = "FARIntAcct"
Public Const FARIntAcctID = "FARIntAcctID"
Public Const FARIntAcctName = "FARIntAcctName"
'Public Const FBank = "FBank"
Public Const FClientID = "FClientID"
Public Const FClientNo = "FClientNo"
Public Const FClientName = "FClientName"
Public Const FDelDate = "FDelDate"
Public Const FDpst = "FDpst"
Public Const FInt = "FInt"
Public Const FIntAcct = "FIntAcct"
Public Const FIntAcctID = "FIntAcctID"
Public Const FInterest = "FInterest"
Public Const FIsAcnt = "FIsAcnt"
Public Const FLastintDate = "FLastIntDate"
Public Const FLoan = "FLoan"
Public Const FODAcct = "FODAcct"
Public Const FODAcctID = "FODAcctID"
Public Const FOpenDate = "FOpenDate"
Public Const FOppAcct = "FOppAcct"
Public Const FOppAcctID = "FOppAcctID"
Public Const FOppAcctName = "FOppAcctName"
Public Const FPayIntMethod = "FPayIntMethod"
Public Const FRate = "FRate"
Public Const FRateDays = "FRateDays"
Public Const FRateType = "FRateType"
Public Const FStatedDpst = "FStatedDpst"
Public Const FSummaryID = "FSummaryID"
Public Const FTZRate = "FTZRate"
Public Const FTZRateDays = "FTZRateDays"
Public Const FTZRateType = "FTZRateType"


'SystemProfiles
Public Const SCProfile = "SC"
'Settlement Center Special System Profiles
Public Const SCLoanPayWithInt = "SCLoanPayWithInt"
Public Const SCIntToLoan = "SCIntToLoan"
Public Const SCDpstPayWithInt = "SCDpstPayWithInt"
Public Const SCIntToDpst = "SCIntToDpst"
Public Const SCIntAutoVch = "SCIntAutoVch"
Public Const SCIntToOneVch = "SCIntToOneVch"
Public Const SCIntMultiSaveNoVch = "SCIntMultiSaveNoVch"
Public Const SCAcctBalStop = "SCAcctBalStop"
Public Const SCIntOneYear = "SCIntOneYear"
Public Const SCYearRateStartDate = "SCYearRateStartDate"
Public Const SCStatIntCalWholeC = "SCStatIntCalWhole"
Public Const SCIntSaveAutoCheck = "SCIntSaveAutoCheck"
Public Const SCConsignSaveAutoCheck = "SCConsignSaveAutoCheck"

Public Function RTYearName() As String
     RTYearName = "年"
End Function
'Public Const RTQuarterName = LoadMKDString("季度",strLanguage)
Public Function RTMonthName() As String
     RTMonthName = "月"
End Function
Public Function RTDayName() As String
    RTDayName = "日"
End Function

