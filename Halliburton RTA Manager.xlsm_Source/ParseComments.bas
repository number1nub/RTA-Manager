Attribute VB_Name = "ParseComments"
' ___________________________________________________________________________________________________
' ***************************************************************************************************
' Module: ParseComments
'
'   This module contains routines and functions that parse through the contents of the
'   comments field in order to find certain information.
'
'   o Priority number is pulled from comments
'   o
'
' About:
'   o Written by:    Rameen Bakhtiary
'   o Last Modified: 2012-11-28
' ___________________________________________________________________________________________________
' ***************************************************************************************************







'____________________________________________________________________________________________________
'====================================================================================================
'   Function: priority
'       Looks for a number followed by a colon at start of a string and returns it. See remarks
'       section for more information on matching.
'
'   Parameters:
'       TheString - Input string to search for priority
'
'   Returns:
'       Returns the value before the colon as a number (double) if found; otherwise, returns 0.
'
'   Remarks:
'       The function will match and return up to 4 characters (as type double) to be returned.
'       This means that a priority number can contain a decimal, as long as the total number of
'       characters doesn't exceed 5 (including the decimal point).
'
'   About:
'       o Written by:       Rameen Bakhtiary
'       o Date Modified:    10/24/2011
'====================================================================================================
    Function priority(TheString As String) As Double
        If TheString = "" Then priority = 0: Exit Function
        tmp = InStr(1, TheString, ":", vbTextCompare)
        If tmp = 0 Then priority = 0: Exit Function
        If tmp < 6 Then priority = Left(TheString, tmp - 1)
    End Function




'____________________________________________________________________________________________________
'====================================================================================================
'   Function: AtAt
'       Checks for '@@' in string. Retrun true if found. False otherwise
'
'   Parameters:
'       TheString - String to search through for '@@'
'
'   Returns:
'       Boolean True if "@@" is found; false otherwise.
'
'   About:
'       o Written by:       Rameen Bakhtiary
'       o Date Modified:    11/28/2011
'====================================================================================================
    Function AtAt(TheString As String)
        tmp = Strings.InStr(1, TheString, "@@", vbTextCompare) ' Find "@@" in the string
        If tmp = 0 Then ' if there are no "@@"
            AtAt = 0
        ElseIf tmp <> 0 Then ' if there is an "@@"
            AtAt = 1
        End If
    End Function
    
    
    
    
'____________________________________________________________________________________________________
'====================================================================================================
'   Function: stdProdLT
'       Get Standard Production Lead Time from the comments by searching for "SPLT=". See the
'       remarks section for more information on string matching.
'
'   Parameters:
'       comTxt - String in which to search for SPLT=
'       valueIfMissing - (Optional) Value to return if "SPLT=" isn't found
'
'   Returns:
'       Returns the SPLT as a number if "SPLT=" is found; Returns [valueIffMissing] if "SPLT=" isn't
'       found. The default value if missing is 0 (empty)
'
'   Remarks:
'       o Matching is not case sensitive; SPLT= or splt= will be matched
'       o Must not have a space before the equal sign. See example below.
'         *CORRECT* - *INCORRECT*
'         SPLT= 85  - SPLT = 85
'         SPLT=85   - SPLT =85
'         splt= 85  - splt = 85
'
'   About:
'       o Written by:       Rameen Bakhtiary
'       o Date Modified:    11/28/2011
'====================================================================================================
    Function stdProdLT(comTxt As String, Optional valueIfMissing = 0) As Double
        comTxt = UCase(comTxt)
        Dim pos As Integer
        
        'Search for "SPLT="
        pos = InStr(1, comTxt, "SPLT=", vbTextCompare)
        
        If pos = 0 Then stdProdLT = valueIfMissing: Exit Function '-- Not Found
        
        'Get tmpStr=
        tmpStr = Trim(Right(comTxt, Len(comTxt) - pos - 4))
        tmp = InStr(1, tmpStr, " ", vbTextCompare)
            If tmp <> 0 Then tmpStr = Left(tmpStr, tmp)
        tmpStr = Trim(tmpStr)
        
        If Not IsNumeric(tmpStr) Or tmpStr = 0 Then
            tmpStr = 0
            stdProdLT = Format(tmpStr, "0.1")
        Else
            stdProdLT = Format(tmpStr, "general number"): Exit Function
        End If
    End Function







