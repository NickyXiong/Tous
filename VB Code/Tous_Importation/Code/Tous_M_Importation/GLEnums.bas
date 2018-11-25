Attribute VB_Name = "GLEnums"
Option Explicit

Public Const SKIPNEGATIVECHECK = 1
Public Const SKIPHIGHLOWCHECK = 2
Public Const SKIPLOCKSTOCK = 4
Public Const ICMOLISTTYPEID = 90  '生产任务单序事簿模板ID
Public Const SUBICMOLISTTYPEID = 620  '委外生产任务单序事簿模板ID

Public Enum ListOptionsEnum
    ebListPrimary
    ebListSecondary
    ebListAll
End Enum

' Object types defined in General Ledger module
Public Enum GLObjectTypeEnum
    ebAccountObject = 0
    ebCurrencyObject = 1
    ebItemClassObject = 2
    ebItemObject = 3
    ebVoucherObject = 4
    ebVoucherGroupObject = 5
    ebPeriodicObject = 6
    ebMeasureUnitObject = 7
    ebLedgerObject = 8
    ebMCLdgObject = 9
    ebSettleTypeObject = 10
    ebAccountItemObject = 11
    ebMeasureUnitGroupObject = 12
    ebContactObject = 13
    ebNoteObject = 15
End Enum
Public Enum GLLedgerObject
    ebGenLdgObject = 1
    ebSubLdgObject = 2
    ebQtyGenLdgObject = 3
    ebQtySubLdgObject = 4
    ebAcctBalObject = 5
    ebTrialObject = 6
    ebItemGenLdgObject = 7
    ebItemSubRptObject = 8
    ebItemBalObject = 9
    ebDailyObject = 10
    ebItemSubAging = 11
    ebItemSumAging = 12
    ebItemCombAging = 13
    ebAcctIntObject = 14
    ebAdjustHistObject = 15
End Enum
Public Enum GLContactObject
    ebContactDueObject = 0
    ebTransStmtObject
    ebAgingObject
End Enum
' Specific access masks
Public Enum GLAccessTypeEnum
    ' Periodic object specific access masks
    ebGLSecAdjustExchangeRate = &H1
    ebGLSecTransferProfitAndLoss = &H2
    ebGLSecClosePeriod = &H4
    
    ' Voucher object specific access masks
    ebGLSecCheckVoucher = &H1
    ebGLSecPostVoucher = &H2
    ebGLSecUnpostVoucher = &H4
End Enum

' GL error codes
Public Enum GLErrorCodes
    ' The range &H80044100-&H800442FF has been allocated to General
    ' Ledger business objects
    EBSGL_E_FIRST = &H80044100
    EBSGL_E_LAST = &H800442FF

    ' Cannot create an account without parent account.
    EBSGL_E_NoParentAccount = &H80044100

    EBSGL_E_VoucherMissingEntries = &H80044101
    EBSGL_E_RequireDetailAccount = &H80044102
    EBSGL_E_CurrencyNotMatch = &H80044103
    EBSGL_E_ItemClassNotMatch = &H80044104
    EBSGL_E_DuplicateItemClass = &H80044105
    EBSGL_E_VoucherDCBothZero = &H80044106
    EBSGL_E_VoucherDCNeitherZero = &H80044107
    EBSGL_E_VoucherNotBalance = &H80044108
    EBSGL_E_PeriodClosed = &H80044109
    EBSGL_E_NotInCurrentPeriod = &H8004410A
    EBSGL_E_VoucherChecked = &H8004410B
    EBSGL_E_VoucherNotCheck = &H8004410C
    EBSGL_E_SameCheckerAndPreparer = &H8004410D
    EBSGL_E_VoucherPosted = &H8004410E
    EBSGL_E_VoucherNotPost = &H8004410F
    EBSGL_E_DuplicateItem = &H80044110
    EBSGL_E_TooManyItemClasses = &H80044111
    EBSGL_E_ItemAmountNotMatch = &H80044112
    EBSGL_E_SpecialVoucher = &H80044113
    EBSGL_E_VoucherPeriodChanged = &H80044114
    EBSGL_E_SummarizedVoucher = &H80044115
    
    EBSGL_E_RateAdjustAccount = &H80044140
    EBSGL_E_EarnAccount = &H80044141
    EBSGL_E_InvalidCoefficient = &H80044150
    EBSGL_E_InvalidUnitGroup = &H80044151
    EBSGL_E_InvalidAccountID2 = &H80044152
    
    EBSGL_E_ItemNotIsAFolder = &H80044160
    EBSGL_E_InValidSeparator = &H80044161
    
    EBSGL_E_NoParentItem = &H80044170
    EBSGL_E_RequireDetailItem = &H80044171
    EBSGL_E_RequireItem = &H80044172
    
    ' ContactUpdate
    EBSGL_E_ContactMissingEntries = &H80044180
    EBSGL_E_ContactDCBothZero = &H80044181
    EBSGL_E_ContactDCNeitherZero = &H80044182
    EBSGL_E_ContactInvalidDate = &H80044182
    
    EBSGL_E_InitializeNotFinished = &H80044200
End Enum

