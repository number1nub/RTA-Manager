Attribute VB_Name = "MyGlobals"
' ___________________________________________________________________________________________________
' ***************************************************************************************************
' Module: GlobalVarS
'   This module contains global settings, constants and variables.
'
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
    Public Const WDPUB As String = "\\corp.halliburton.com\team\WD\Public\"
    '___________________________________
    '       RTA SHEET MAIN PUBLIC FOLDER
    '
    Public Const PUB_DIR  As String = WDPUB & "RTA Manager\"
    '____________________________________________________________________
    '       LOCATION WHERE RTALOAD FILES ARE BACKED UP/SAVED ON EACH LOAD
    '
    Public Const BACKUP_DIR  As String = PUB_DIR & "RTALoad Archive\"
    '______________________________________________________________
    '       DEVELOPER FOLDER PATH - If the workbook is in this path
    '       then workbook save restrictions are ignored
    '
    Public Const DEV_PATH As String = "C:\Dropbox\Halliburton RTA Manager"
    '________________________________
    '       PATH TO UPDATE CHECK FILE
    '
    Public Const UPDATE_PATH As String = WDPUB & "Rameen Bakhtiary\wdRTApush.exe"

    '____________________________________________________
    '       DEFAULT SHEET PASSWORD (USED FOR SWITCH MODE)
    '
    Public Const DEFAULT_PW As String = "wdr74!"

    
    '_________________________________
    '       SHEET   GLOBAL   VARIABLES
    '
    Public myPath As String
    Public MyDocs As String
    Public WinUname As String
    Public FullName As String
        
    
    
' ===================================================================================================
' I N I T I AL I Z E   G L O B A L    S H E E T    V A R I A B L E S
' ===================================================================================================

Public Sub initializeGlobals()


    '________________________
    '       PATH TO THIS FILE
    '
    myPath = ThisWorkbook.Path
    
    '____________________________________________________________________
    '       USER'S LOGON USERNAME & USER'S FULL FIRST/LAST NAME (IF FOUND)
    '
    WinUname = UserName
    FullName = userFullName
    
    '__________________________
    '       MY DOCUMENTS FOLDER
    '
    MyDocs = "C:\Documents and Settings\" & WinUname & "\My Documents\"
End Sub
    
    
    
    
    
    
    
    
    
    
    
    
    
    
