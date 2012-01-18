Attribute VB_Name = "MyGlobals"
' ___________________________________________________________________________________________________
' ***************************************************************************************************
' Module: GlobalVars
'
'   This module contains global settings, constants and variables.
'
' About:
'   o Last Modified: 2012-01-16
' ___________________________________________________________________________________________________
' ***************************************************************************************************


'===================================================================================================
'                                   F O L D E R    P A T H S
'===================================================================================================
  
    '________________________________
    '       ENGINEERING PUBLIC FOLDER
    '
    Public Const EPUB  As String = "\\corp.halliburton.com\team\WD\Business Development and Technology\General\Engineering Public\"
    '_______________________
    '       WD PUBLIC FOLDER
    '
    Public Const WDPUB As String = "\\corp.halliburton.com\team\WD\"
    '___________________________________
    '       RTA SHEET MAIN PUBLIC FOLDER
    '
    Public Const PUB_DIR  As String = WDPUB & "RTA Manager\"
    '____________________________________________________________________
    '       LOCATION WHERE RTALOAD FILES ARE BACKED UP/SAVED ON EACH LOAD
    '
    Public Const BACKUP_DIR  As String = PUB_DIR & "RTALoad Archive\"
    '________________________________________________________________________________
    '       DEVELOPER FOLDER PATH - If the workbook is in this path
    '       then workbook save restrictions are ignored
    '
    Public Const DEV_PATH As String = "C:\Dropbox\Halliburton RTA Manager"


'===================================================================================================
'                    R T A   S H E E T   H E L P E R   F I L E   P A T H S
'===================================================================================================

    '________________________________
    '       PATH TO UPDATE CHECK FILE
    '
    Public Const UPDATE_PATH As String = PUB_DIR & "RTA Sheet Update.exe"
    
    
'===================================================================================================
'                               O T H E R   C O N S T A N T S
'===================================================================================================

    '____________________________________________________
    '       DEFAULT SHEET PASSWORD (USED FOR SWITCH MODE)
    '
    Public Const DEFAULT_PW As String = "wdr74!"





