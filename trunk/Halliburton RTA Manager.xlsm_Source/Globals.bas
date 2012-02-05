Attribute VB_Name = "Globals"
' ___________________________________________________________________________________________________
' ***************************************************************************************************
' Module: GlobalVarS
'   This module contains global settings, constants and variables.
'
' ***************************************************************************************************




'===================================================================================================
'                              C O N S T A N T     F O L D E R    P A T H S
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



'===================================================================================================
'                             S H E E T      S E T T I N G S
'===================================================================================================

    '____________________________________________________
    '       DEFAULT SHEET PASSWORD (USED FOR SWITCH MODE)
    '
        Public Const DEFAULT_PW As String = "wdr74!"

    
    '_______________________________________________________________________
    '       LIST OF EMAILS TO 'CC:' WHEN THE EMAIL REQUESTOR BUTTON PRESSED
    '       (List as one string, separate with semi-colons (;) )
    '
        Public Const RtaLiasonEmail As String = "Dana.Moe@halliburton.com"
        
    '_________________________________________________________________________
    '       FORMAT OF THE COMMENT LINE START FOR LIASON/COMMENTS SECTION
    '
        Public Const LiasonCommentDateFormat As String = "M/d/yyyy"
        Public Const LiasonCommentInitialsFormat As String = " DM, "
        
    '_________________________________________________________________________
    '       FORMAT OF THE COMMENT LINE START FOR DESCRIPTION / WEEKLY MEETING
    '
        Public Const WeeklyMeetingDateFormat As String = "yyyy-MM-dd"
        Public Const WeeklyMeetingInitialsFormat As String = ", WM: "
        
    '_____________________________________________
    '       DETERMINES IF USER CAN WRITE COMMENTS
    '       (For use in liason mode)
    '
        Public CanComment As Boolean        ''' Not yet implemented






'===================================================================================================
'        U S E R   S P E C I F I C    P A T H S - A S S I G N   O N   S T A R T
'===================================================================================================
    
    '________________________________________
    '       PATHS TO THE INCLUDED EXE FILES
    '
        Public CMDlinePath As String
        Public SheetToolsPath As String
    
    
    '_________________________________________
    '       USER-SPECIFIC   GLOBAL   VARIABLES
    '
        Public pubInit As Boolean       '- Can be checked to see if user vars are assigned
        Public myPath As String         '- Full path to this sheet
        Public MyDocs As String         '- Path to user's my documents folder
        Public WinUname As String       '- Windows username of user
        Public FullName As String       '- User's full name (attempt to pull from properties)
        Public userInitials As String   '- User's initials derived from fullName
    
    '_____________________________________________________________
    '       RTA SPECIFIC VALUES - ASSIGNED WHEN AN RTA IS SELECTED
    '
        Public RTAselected As Boolean   '- Used to determine if the public RTA
                                        '  vars are currently accurate
        Public thisRow As Integer       '- Current active row
        
        Public thisRta As String
        Public thisRtaLong As String
        Public thisClass As String
        Public thisClassLong As String
        Public thisDescription As String
        Public thisComment As String
        Public thisAssignedto As String
        Public thisDept As String
        Public thisTRDD As String
        Public thisLabOffice As String
        Public thisType As String
        Public thisCode As String
        Public thisRequestor As String
        Public thisRequestorEmail As String
        Public thisSubmitter As String
        Public thisState As String
        
      
        
' %&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%
    
' ===================================================================================================
' I N I T I AL I Z E   G L O B A L    S H E E T    V A R I A B L E S
' ===================================================================================================

Public Sub initGlobals()

    '________________________
    '       PATH TO THIS FILE
    '
    myPath = ThisWorkbook.Path
    
    
    '___________________________________
    '       RTA SHEET TOOLS.EXE FILEPATH
    '
    SheetToolsPath = myPath & "\Include\RTA Sheet Tools.exe"
    
    '_____________________________________
    '       CMDline_Functions.exe FILEPATH
    '
    CMDlinePath = myPath & "\Include\CMDline_Functions.exe"
    
    
    '____________________________________________________________________
    '       USER'S LOGON USERNAME & USER'S FULL FIRST/LAST NAME (IF FOUND)
    '
    WinUname = UserName
    FullName = userFullName
    
    '________________________
    '       GET USER INITIALS
    '
    splitNames = Split(FullName, " ", , vbTextCompare)
    userInitials = UCase(Left(Trim(splitNames(0)), 1) & Left(Trim(splitNames(1)), 1))
    
    '__________________________
    '       MY DOCUMENTS FOLDER
    '
    MyDocs = "C:\Documents and Settings\" & WinUname & "\My Documents\"
    
    
    pubInit = (myPath <> "")
End Sub





''' ===================================================================================================
''' I N I T I AL I Z E   T H E   P E R - R T A   V A R I A B L E S
''' ---------------------------------------------------------------------------------------------------
'''   o thisRow           - Current row number
'''   o thisRTA           - Current RTA number
'''   o thisRTAlong       - RTA number in R00000XXXXXX format
'''   o thisClass         - Current RTAs class
'''   o thisClassLong     - Full class name as in CWI
'''   o thisDescription   - Current RTAs description
'''   o thisComment       - Current RTAs comments
'''   o thisAssignedto    - Current RTAs assigned to
'''   o thisDept          - Current RTAs current status
'''   o thisTRDD          - Current RTAs revised due date
'''   o thisLabOffice     - Current RTAs lab office
'''   o thisType          - Current RTAs type
'''   o thisCode          - Current RTAs code
'''   o thisRequestor     - Current RTAs requestor name
'''   o thisRequestorEmail- Current requestor'''s email
'''   o thisSUbmitter     - Current RTAs submittrer name
'''   o thisState         - Current RTAs state
''' ===================================================================================================
Public Sub getCurrent()
    
    thisRow = ActiveCell.Row
            
    If thisRow < 6 Then RTAselected = False: Exit Sub
                        
    thisRta = ActiveSheet.Cells(thisRow, getCol("RTA"))
    thisRtaLong = "R00000" & thisRta
    thisClass = Cells(thisRow, getCol("class"))
    Select Case thisClass
        Case "A"
        thisClassLong = "A=Minimal Processing Time"
        Case "B"
        thisClassLong = "B=Medium Processing Time"
        Case "C"
        thisClassLong = "C=Technology Negotiated Processing Time"
        Case "D"
        thisClassLong = "D=Technology Development Engineering"
    End Select
    thisDescription = Cells(thisRow, getCol("description"))
    thisComment = Cells(thisRow, getCol("comments"))
    thisAssignedto = Cells(thisRow, getCol("assigned to"))
    thisDept = Cells(thisRow, getCol("current status"))
    thisTRDD = Cells(thisRow, getCol("revised due date"))
    thisLabOffice = Cells(thisRow, getCol("lab office"))
    thisType = Cells(thisRow, getCol("type"))
    thisCode = Cells(thisRow, getCol("code"))
    thisRequestor = Cells(thisRow, getCol("requestor name"))
    thisRequestorEmail = Cells(thisRow, getCol("requestor email"))
    thisSubmitter = Cells(thisRow, getCol("requestor name"))
    thisState = Cells(thisRow, getCol("state"))
    
End Sub
    
    
    
    
    
    
    
    
    
    
    
    
