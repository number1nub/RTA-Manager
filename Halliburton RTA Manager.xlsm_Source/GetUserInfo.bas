Attribute VB_Name = "GetUserInfo"
'******** Code Start ********
'This code was originally written by Dev Ashish.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Dev Ashish
'
Private Type USER_INFO_2
    usri2_name As Long
    usri2_password  As Long  ' Null, only settable
    usri2_password_age  As Long
    usri2_priv  As Long
    usri2_home_dir  As Long
    usri2_comment  As Long
    usri2_flags  As Long
    usri2_script_path  As Long
    usri2_auth_flags  As Long
    usri2_full_name As Long
    usri2_usr_comment  As Long
    usri2_parms  As Long
    usri2_workstations  As Long
    usri2_last_logon  As Long
    usri2_last_logoff  As Long
    usri2_acct_expires  As Long
    usri2_max_storage  As Long
    usri2_units_per_week  As Long
    usri2_logon_hours  As Long
    usri2_bad_pw_count  As Long
    usri2_num_logons  As Long
    usri2_logon_server  As Long
    usri2_country_code  As Long
    usri2_code_page  As Long
End Type
 
 
 
 
Private Declare Function apiNetGetDCName Lib "netapi32.dll" Alias "NetGetDCName" (ByVal servername As Long, ByVal DomainName As Long, bufptr As Long) As Long
 
' function frees the memory that the NetApiBufferAllocate
' function allocates.
Private Declare Function apiNetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (ByVal buffer As Long) As Long
 
' Retrieves the length of the specified wide string.
Private Declare Function apilstrlenW Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
 
Private Declare Function apiNetUserGetInfo Lib "netapi32.dll" Alias "NetUserGetInfo" (servername As Any, UserName As Any, ByVal level As Long, bufptr As Long) As Long
 
' moves memory either forward or backward, aligned or unaligned,
' in 4-byte blocks, followed by any remaining bytes
Private Declare Sub sapiCopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
 
Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
 
Private Const MAXCOMMENTSZ = 256
Private Const NERR_SUCCESS = 0
Private Const ERROR_MORE_DATA = 234&
Private Const MAX_CHUNK = 25
Private Const ERROR_SUCCESS = 0&
 
 
 
 
 
' ___________________________________________________________________________________________________
' ===================================================================================================
' Function: userFullName
'
'   -
'
' Parameters:
'   -
'
' Last Modified: 2012-01-17
' ___________________________________________________________________________________________________
' ===================================================================================================
Function userFullName(Optional strUserName As String) As String
'
    ' Returns the full name for a given UserID
    '   NT/2000 only
    ' Omitting the strUserName argument will try and
    ' retrieve the full name for the currently logged on user
    '
    On Error GoTo ErrHandler
    Dim pBuf As Long
    Dim dwRec As Long
    Dim pTmp As USER_INFO_2
    Dim abytPDCName() As Byte
    Dim abytUserName() As Byte
    Dim lngRet As Long
    Dim i As Long
     
    ' Unicode
    abytPDCName = fGetDCName() & vbNullChar
    If (Len(strUserName) = 0) Then strUserName = UserName()
    abytUserName = strUserName & vbNullChar
 
    ' Level 2
    lngRet = apiNetUserGetInfo(abytPDCName(0), abytUserName(0), 2, pBuf)
    If (lngRet = ERROR_SUCCESS) Then
        Call sapiCopyMem(pTmp, ByVal pBuf, Len(pTmp))
        userFullName = fStrFromPtrW(pTmp.usri2_full_name)
    End If
 
    Call apiNetAPIBufferFree(pBuf)
ExitHere:
    Exit Function
ErrHandler:
    userFullName = vbNullString
    Resume ExitHere
End Function
 
 
 
 
' ___________________________________________________________________________________________________
' ===================================================================================================
' Function: UserName
'   Returns the user's network logon username.
'
' Last Modified: 2012-01-17
' ___________________________________________________________________________________________________
' ===================================================================================================
Function UserName() As String
    '----------------------------------------------------------
    ' This code was originally written by Dev Ashish.
    ' It is not to be altered or distributed,
    ' except as part of an application.
    ' You are free to use it in any application,
    ' provided the copyright notice is left unchanged.
    '
    ' Code Courtesy of
    ' Dev Ashish
    '

    Dim lngLen As Long, lngX As Long
    Dim strUserName As String
    strUserName = String$(254, 0)
    lngLen = 255
    lngX = apiGetUserName(strUserName, lngLen)
    If (lngX > 0) Then
        UserName = Left$(strUserName, lngLen - 1)
    Else
        UserName = vbNullString
    End If
End Function


 
 
 
' ___________________________________________________________________________________________________
' ===================================================================================================
' Function: fGetDCName
'
'   -
'
' Parameters:
'   -
'
' Last Modified: 2012-01-17
' ___________________________________________________________________________________________________
' ===================================================================================================
Function fGetDCName() As String
Dim pTmp As Long
Dim lngRet As Long
Dim abytBuf() As Byte
 
    lngRet = apiNetGetDCName(0, 0, pTmp)
    If lngRet = NERR_SUCCESS Then
        fGetDCName = fStrFromPtrW(pTmp)
    End If
    Call apiNetAPIBufferFree(pTmp)
End Function
 
 
 
' ___________________________________________________________________________________________________
' ===================================================================================================
' Function: fStrFromPtrW
'
'   -
'
' Parameters:
'   -
'
' Last Modified: 2012-01-17
' ___________________________________________________________________________________________________
' ===================================================================================================
Private Function fStrFromPtrW(pBuf As Long) As String
Dim lngLen As Long
Dim abytBuf() As Byte
 
    ' Get the length of the string at the memory location
    lngLen = apilstrlenW(pBuf) * 2
    ' if it's not a ZLS
    If lngLen Then
        ReDim abytBuf(lngLen)
        ' then copy the memory contents
        ' into a temp buffer
        Call sapiCopyMem(abytBuf(0), ByVal pBuf, lngLen)
        ' return the buffer
        fStrFromPtrW = abytBuf
    End If
End Function
' ******** Code End *********
