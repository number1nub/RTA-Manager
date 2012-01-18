Attribute VB_Name = "PasswordProtect"
'***************************************************************************************************
'   Title: PASSWORD_PROTECT
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   Functions and routines for protecting certain features from unauthorized users.
'
'   Group: About
'       General script information
'       - *Written by:* Rameen Bakhtiary
'       - *Created on:* 2012-01-12
'___________________________________________________________________________________________________
'***************************************************************************************************

Public OK As Boolean       'Used to check for valid password entry
Public sMyPassWord As String        'Password to unlock certain sheet features





'____________________________________________________________________________________________________
'====================================================================================================
'   Function: GetPassword
'       Displays a prompt to enter a password and sets the value of GetPassWord (T/F)
'       depending on whether or not the correct PW was entered.
'
'   Parameters:
'       title   -   Title of the password prompt window
'       tmpPW   -   *(Optional)* The correct password to wait for. If left blank then will
'                   defaults to the constant 'DEFAULT_PW' defined in the MyGlobals module.
'
'====================================================================================================
Function GetPassword(title As String, Optional tmpPW As String = DEFAULT_PW)
    sMyPassWord = tmpPW
    protect.Caption = title
    protect.Show
    GetPassword = OK
End Function

