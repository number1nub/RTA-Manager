VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RTA_Details 
   Caption         =   "Modify / View RTA Details"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   13395
   OleObjectBlob   =   "RTA_Details.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RTA_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************
'   Title: RTA_Details GUI
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   A simple GUI for viewing and editing RTA details
'
'   Group: About
'       General script information
'       - *Written by:* Rameen Bakhtiary
'       - *Created on:* 2012-01-12
'___________________________________________________________________________________________________
'***************************************************************************************************

Public savedForm As Boolean






'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: UserForm_Initialize
'       Code executed every time after the RTA Details form is called and before the GUI is shown.
'
'====================================================================================================
Private Sub UserForm_Initialize()
    
    '________________________________
    '       SET THE PUBLIC RTA VALUES
    '
        getCurrent
        
        
    '___________________________
    '       FILL THE RTA INFO IN
    '
    rtaNUm.Caption = "RTA " & thisRta
    class = thisClass
    desc = thisDescription
    comments = thisComment
    assignedTo = thisAssignedto
    Department = thisDept
    techrevdate = thisTRDD
    lab = thisLabOffice
    rtaType = thisType
    rtacode = thisCode
    requestor = thisRequestor
    state = thisState


    '________________________________
    '       SET THE LAB OFFICE PREFIX
    '
    Select Case lab
    Case "WD1"
        prefix = "fc"
    Case "WD2"
        prefix = "di"
    Case "WD3"
        prefix = "pm"
    Case "WD4"
        prefix = "fc"
    Case "WD5"
        prefix = "S"
    Case Else
        prefix = ""
    End Select
    
    '______________________________________________________________________________
    '       FILL THE 'ASSIGNED TO' COMBOBOX WITH ALL NAMES FROM THE RTAS LAB OFFICE
    '
    On Error Resume Next
    tmpary = Application.Range("Name" & prefix)
    For Each v In tmpary
        assignedTo.AddItem (v)
    Next
    
    '_____________________________________________
    '       ENABLE EDITABLE CONTROLS BASED ON MODE
    '
    If Application.Range("sheetViewMode") = "EDIT" Then
        assignedTo.Enabled = True
        class.Enabled = True
        techrevdate.Enabled = True
        Department.Enabled = True
    End If
    
    savedForm = True
    uploadToCWI.Enabled = False
    desc.SetFocus
    desc.SelStart = 0
End Sub





 

'
'===================================================================================================
'   Sub: emailSubmitter_Click
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   Open a new email to the RTA requestor
'
'===================================================================================================
Private Sub emailRequestor_Click()
    
    '______________________________
    '   GET FIRST NAME OF REQUESTOR
    '
    Dim firstName As String, name As String
    name = Trim(Split(thisRequestor, " ", , vbTextCompare))
    firstName = name(0)
    
    
    '__________________________________________________
    '   MAKE SURE THERE ARE NO CHANGES TO SAVE B4
    '   CLOSING THE GUI & OPENING AN EMAIL
    '
    If Application.Range("sheetviewmode") = "EDIT" And savedForm = False Then
        Call uploadToCWI_Click
    End If
    
    '________________________________________________________________
    '    OPEN AN EMAIL TO THE REQUESTOR W/THE RTA IN THE SUBJECT &
    '    THE 'RTALIASONEMAIL' (SET IN GLOBAL SETTINGS) AS A CC
    '
    emailstring = "mailto:" & thisRequestorEmail & _
                "?cc=" & RtaLiasonEmail & _
                "&subject=RTA " & thisRta & _
                "&body=" & firstName & ", " & Chr(10) & Chr(10)

    ActiveWorkbook.FollowHyperlink emailstring
        
End Sub






' ___________________________________________________________________________________________________
' ===================================================================================================
' Sub: resetME_Click  -  RESET BUTTON
'
' Last Modified: 2012-02-02
' ___________________________________________________________________________________________________
' ===================================================================================================
Private Sub resetME_Click()
    Call UserForm_Initialize
End Sub



'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: changeMade
'       Make note when any change is made to the GUI and mark it as NOT savedForm. Notify the user
'       if they close without saving
'
' Remarks:
'       ONLY VALID WHEN IN EDIT MODE
'
'====================================================================================================
Private Sub assignedTo_Change()
    Call changeMade
End Sub
Private Sub class_Change()
    Call changeMade
End Sub
Private Sub comments_Change()
    Call changeMade
End Sub
Private Sub Department_Change()
    Call changeMade
End Sub
Private Sub desc_Change()
    Call changeMade
End Sub
Private Sub rtatype_Change()
    Call changeMade
End Sub
Private Sub techrevdate_Change()
    Call changeMade
End Sub
Sub changeMade()
    savedForm = False
    If Application.Range("sheetViewMode") <> "PMT" Then uploadToCWI.Enabled = True
End Sub







' ===================================================
'   CWI  VIEW  BUTTONS
' ===================================================

Private Sub viewRTA_Click()
    Call openCWIpage("v")
End Sub

Private Sub history_Click()
    Call openCWIpage("h")
End Sub

Private Sub structure_Click()
    Call openCWIpage("s")
End Sub


Private Sub modRta_Click()
    Call openCWIpage
End Sub

Private Sub printRta_Click()
    Call openCWIpage("p")
End Sub



'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: DOUBLE-CLICK rtaNum
'       Open the RTAs page in CWI or search it.
'
'   Remarks:
'       Uses the external file: CMDline_Functions.exe
'====================================================================================================
Private Sub rtaNUm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    Call openCWIpage
End Sub
Private Sub rtaType_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    Call openCWIpage
End Sub









' ___________________________________________________________________________________________________
' ===================================================================================================
' Sub: openCWIpage
'   Uses the external file 'CMDline_Functions.exe' to open the current RTA in the specified CWI
'   view.
'
' Parameters:
'   view    -   (Optional) A 1 to 3 letter identifier that determines what CWI view to open
'               the RTA in. See the list below.
'
' View modes:
'   n           -   Opens in Navigate view
'   sig         -   Opens in View Approvals/Promote view
'   wu          -   Opens in Where Used view
'   m           -   Opens in Modify view
'   rta         -   Opens in CreateModify RTA view
'   h           -   Opens history of the part
'   p           -   Opens print view
' Blank or any other  -   Open in StructureManagement
'
'
' Last Modified: 2012-02-04
' ___________________________________________________________________________________________________
' ===================================================================================================
Sub openCWIpage(Optional view As String = "rta")
    
    '_____________________________________
    '       ENSURE THAT GLOBALS ARE LOADED
    '
    If Not pubInit Then initGlobals
    

    '____________________________________________________________________
    ' THIS IS GETTING A BIT (VERY) UNNECESSARY SINCE SHEET DOESN'T ALLOW
    ' ENTRY IF THIS ISN'T TRUE...
    '
    If Dir(progP) = "" Then
        Call MsgBox("Uh oh... An important file couldn't be found:  CMDline_Functions.exe" & _
            vbCrLf & "" & vbCrLf & "Without it, you cannot open RTAs directly in CWI. This file should be located" & _
            Chr(10) & "in the Include folder within this worksheets directory. " & vbCrLf & vbCrLf & _
            "Running the installer file should solve this issue", vbCritical Or vbSystemModal, "--=[ WD RTA Sheet ]=--")
        Exit Sub
    End If
    
    '________________________________________________________
    ' IN EDIT MODE... CHECK FOR CHANGES & SAVE B4 CLOSING GUI
    '
    If Application.Range("sheetviewmode") = "Edit" Then
        If savedForm = False Then Call uploadToCWI_Click
    End If
    
    
    '_____________________________
    ' CLOSE THE GUI AND OPEN IE
    '
    Unload Me
    
    Call CMDline("""" & progP & """ " & thisRta & " " & view, vbNormalFocus)
    Exit Sub
End Sub





'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: closeME_Click
'       Check that all changes have been savedForm and close the GUI
'
'====================================================================================================
Private Sub closeME_Click()
    If savedForm = False And Application.Range("sheetviewmode") <> "PMT" Then
        ans = MsgBox("" & vbCrLf & "You made changes to this RTA that have not been savedForm." & vbCrLf & "" & vbCrLf & "DISCARD CHANGES?", _
                vbYesNo Or vbExclamation Or vbSystemModal Or vbDefaultButton1, "     DISCARD CHANGES??")
        If ans = vbNo Then desc.SetFocus: Exit Sub
    End If
    Unload Me
End Sub




'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: uploadToCWI_Click
'
'       Writes the information currently shown on the GUI onto the RTAimport sheet
'       formatted so that it can be loaded into CWI when finished using the Modify objects
'       from Excel tool in CWI.
'
'====================================================================================================
Private Sub uploadToCWI_Click()

    On Error GoTo err1

    Dim rtasht As Worksheet
    Set rtasht = ThisWorkbook.Sheets("RTAimport")
    Application.ScreenUpdating = False
    rtasht.Visible = xlSheetVisible
    rtasht.Select
    
    
    ' Find the first open cell on RTAimport sheet or find the same
    ' RTA already on the sheet to overwrite
    '===============================================================
    r = 1
    While Cells(r, 1) <> ""
        If Cells(r, 2) = thisRtaLong Then GoTo overwrite
        r = r + 1
    Wend
    
    
overwrite:
    Range("a" & r) = "Rta"
    Range("b" & r) = thisRtaLong
    
    
    ' Remove carriage returns & remove multiple blank lines
    ' from comments and description
    '=======================================================
    Range("c" & r) = Replace(Replace(desc.Text, Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10), Chr(10)), Chr(13), "")
    Range("d" & r) = Replace(Replace(comments.Text, Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10), Chr(10)), Chr(13), "")
    
    
    
    '   RTA CLASS - get full text & insert
    '===================================================
    Select Case class
    Case "A"
        fullc = "A=Minimal Processing Time"
    Case "B"
        fullc = "B=Medium Processing Time"
    Case "C"
        fullc = "C=Technology Negotiated Processing Time"
    Case "D"
        fullc = "D=Technology Development Engineering"
    End Select

    Range("e" & r) = fullc
    
    
    '   ASSIGNED TO - Insert
    '===================================================
    Range("f" & r) = assignedTo
    
    
    
    '   ASSIGNED TO DEPARTMENT - insert
    '===================================================
    Range("g" & r) = Department
    
    
    
    '   TECH REV DATE - convert to string value & insert
    '===================================================
    Range("h" & r) = techrevdate


    '===================================================
    '   Create and save a copy of the RTAimport sheet
    '===================================================
        Application.DisplayAlerts = False
                    

        ' SAVE RTAload.xlsx in My Documents
        '========================================
        rtasht.Select
        rtasht.Copy
        
        ActiveWorkbook.SaveAs Filename:=MyDocs & "rtaLoad.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
        
        
        'Close the new workbook just created
        '=======================================
        ActiveWorkbook.Close
    
    
    
    Application.DisplayAlerts = True
    rtasht.Visible = xlSheetHidden
    
    Sheets("RTA Manager").Select
    arw = ActiveCell.Row
    
    Cells(arw, getCol("class")) = class
    Cells(arw, getCol("Description")) = Replace(desc.Text, Chr(13), "")
    Cells(arw, getCol("Comments")) = Replace(Replace(comments.Text, Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10), Chr(10)), Chr(13), "")
    Cells(arw, getCol("Assigned To")) = assignedTo
    Cells(arw, getCol("Current Status")) = Department
    Cells(arw, getCol("Revised Due Date")) = techrevdate
    Application.ScreenUpdating = True
    
    Unload Me
    Exit Sub
    
    
    
    '===================================================
    '          ERROR   HANDLER
    '
    '   Make sure all background sheets get hidden in
    '   case of an error.
    '===================================================
err1:
    Workbooks("Halliburton RTA Manager.xlsm").Sheet2.Visible = xlSheetHidden
    Workbooks("Halliburton RTA Manager.xlsm").Sheet3.Visible = xlSheetHidden
    Workbooks("Halliburton RTA Manager.xlsm").Sheet1.Select
    Range("A1").Select
    

    
End Sub




' ___________________________________________________________________________________________________
' ===================================================================================================
' Sub: wmComment_Click
'       Inserts the formatted weekly meeting comment start (YYYY-MM-DD, WM: ) at the bottom of the
'       description
'
' Last Modified:
'       2012-02-01
' ___________________________________________________________________________________________________
' ===================================================================================================
Private Sub wmComment_Click()
    
    If Application.Range("sheetviewmode") = "PMT" Then Call MsgBox( _
        "You must be in Edit mode to insert a weekly meeting comment!", _
        vbExclamation Or vbSystemModal, "Insert Weekly Meeting Comments"): Exit Sub

    insertTxt = Chr(10) & Chr(10) & Format(Now(), WeeklyMeetingDateFormat) & WeeklyMeetingInitialsFormat
    desc = desc & insertTxt
    desc.SelStart = 5000
    
End Sub



' ___________________________________________________________________________________________________
' ===================================================================================================
'   Sub: liasonComment_Click
'
' Last Modified:
'       2012-2-4
' ___________________________________________________________________________________________________
' ===================================================================================================
Private Sub liasonComment_Click()

    '___________________________________________________
    '   NOT IN EDIT MODE - OR - HAS 'NO COMMENT' MODE ON
    '   ("No comment" mode not yet implemented)
    '
    If Application.Range("sheetviewmode") = "PMT" Then Call MsgBox( _
        "You must be in Edit mode to insert a comment!", _
        vbExclamation Or vbSystemModal, "Insert Comments"): Exit Sub
    

    insertTxt = Chr(10) & Chr(10) & Format(Now(), LiasonCommentDateFormat) & LiasonCommentInitialsFormat
    desc = desc & insertTxt
    desc.SelStart = 5000
End Sub










