Attribute VB_Name = "SysEnums"
Option Explicit

Public Enum EBOSystemCheckContants
    EBOSys_OK = 0
    EBOSys_NoDataSource = 1
    EBOSys_CanNotOpenDataSource = 2
    EBOSys_NotEBODatabase = 3
    EBOSys_MissTables = 4
End Enum
' Base error codes
Public Enum BaseErrorCodes
    ' The range &H80044000-&H800440FF has been allocated to base
    ' business objects
    EBSBASE_E_FIRST = &H80044000
    EBSBASE_E_LAST = &H800440FF
    
    '-------------------------------------------------------------------
    ' System profile related error codes
    '-------------------------------------------------------------------
    
    ' Invalid profile category or key name.
    EBSBASE_E_BadProfileKey = &H80044000
    
    ' Cannot modify or delete readonly profile value.
    EBSBASE_E_ProfileReadonly = &H80044001
    
    '-------------------------------------------------------------------
    ' Security related error codes
    '-------------------------------------------------------------------
    
    ' The specified user is forbidden
    EBSSEC_E_UserIsForbidden = &H80044009
    
    ' Access is denied
    EBSSEC_E_AccessDenied = &H80044010
    
    ' The user name or group name is invalid.
    EBSSEC_E_InvalidUsername = &H80044011

    ' The specified domain account not found or not a valid type.
    EBSSEC_E_InvalidDomainAccount = &H80044012
    
    ' The specified domain account already has a difference user name.
    EBSSEC_E_DomainAccountExists = &H80044013
    
    ' The specified user already exists.
    EBSSEC_E_UserExists = &H80044014

    ' The specified user doesn not exist.
    EBSSEC_E_NoSuchUser = &H80044015

    
    ' The specified group already exists.
    EBSSEC_E_GroupExists = &H80044016
    
    ' The specified group does not exist.
    EBSSEC_E_NoSuchGroup = &H80044017
    
    ' Either the specified user is already a member of the specified
    ' group, or the specified group cannot be deleted because it
    ' contains a member.
    EBSSEC_E_MemberInGroup = &H80044018
    
    ' The specified user is not a member of the specified group.
    EBSSEC_E_MemberNotInGroup = &H80044019
    
    ' The last remaining administration user cannot be deleted.
    EBSSEC_E_LastAdmin = &H8004401A
    
    ' Cannot perform this operation on this built-in special group.
    EBSSEC_E_SpecialGroup = &H8004401B
    
    ' Cannot perform this operation on this built-in special user.
    EBSSEC_E_SpecialUser = &H8004401C
    
    ' To many groups have been specified.
    EBSSEC_E_TooManyGroups = &H8004401D
    
    ' Too many users have been specified.
    EBSSEC_E_TooManyUsers = &H8004401E
    
    ' A specified privilege does not exists.
    EBSSEC_E_NoSuchPrivilege = &H8004401F
    
    ' A required privilege is not held by the client.
    EBSSEC_E_PrivilegeNotHeld = &H80044020

    ' Each user belong to 'Users' group
    EBSSEC_E_CanNotRemoveUserFromUsers = &H80044021
    '-------------------------------------------------------------------
    ' Generic business object errors
    '-------------------------------------------------------------------
    
    ' The specified object already exists.
    EBS_E_ObjectExists = &H80044040
    
    ' The specified object does not exists.
    EBS_E_ObjectNotFound = &H80044041
    
    ' Cannot perform this operation on this built-in special object.
    EBS_E_SpecialObject = &H80044042
    
    ' The required property is missing.
    EBS_E_MissingRequiredData = &H80044043
    
    ' Data type mismatch.
    EBS_E_TypeMismatch = &H80044044
    
    ' Data value overflow.
    EBS_E_DataOverflow = &H80044045
    
    ' Literal data is too long.
    EBS_E_DataTooLong = &H80044046
    
    ' Invalid Data value.
    EBS_E_InvalidData = &H80044047
    
    ' Object is in use.
    EBS_E_ObjectInUse = &H80044048
    
    ' Object has one or more children.
    EBS_E_ObjectHasChildren = &H80044049
    
End Enum
'  The following are masks for the standard access types
Public Enum StandardAccessTypeEnum
    ebSecAllAccess = &H80000000
    
    ebSecDelete = &H10000
    ebSecCreate = &H20000
    ebSecReadControl = &H40000
    ebSecWriteControl = &H80000
    ebSecListData = &H100000
    ebSecReadData = &H200000
    ebSecWriteData = &H400000
    
    ebSecStandardAll = &H7F0000
    ebSecGenericRead = ebSecReadControl + ebSecListData + ebSecReadData
    ebSecGenericWrite = ebSecReadControl + ebSecWriteData
    
    ebSecSpecificAll = &HFFFF&
End Enum


