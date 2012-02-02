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

Public saved As Boolean




'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: UserForm_Initialize
'       Code executed every time after the RTA Details form is called and before the GUI is shown.
'
'====================================================================================================
Private Sub UserForm_Initialize()
    Dim crw As Integer
    crw = ActiveCell.Row
    rtaNUm.Caption = "RTA " & ActiveCell.value
    class = Cells(crw, getCol("Class"))
    desc = Cells(crw, getCol("Description"))
    comments = Cells(crw, getCol("Comments"))
    assignedTo = Cells(crw, getCol("Assigned To"))
    Department = Cells(crw, getCol("Current Status"))
    techrevdate = Cells(crw, getCol("Revised Due Date"))
    lab = Cells(crw, getCol("Lab Office"))
    rtaType = Cells(crw, getCol("Type"))
    rtacode = Cells(crw, getCol("Code"))
    
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
    '=== Fill the 'Assigned To' combobox with all names from the RTAs Lab Office
    On Error Resume Next
    tmpary = Application.Range("Name" & prefix)
    For Each v In tmpary
        assignedTo.AddItem (v)
    Next
    
    If Application.Range("sheetViewMode") <> "PMT" Then pmtHide.Enabled = True
    
    saved = True
    desc.SelStart = 0
    uploadToCWI.Enabled = False

End Sub







'
'===================================================================================================
'   Sub: emailSubmitter_Click
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   -
'
'   Parameters:
'       -
'
'   Group: About
'       - *Written by:* Rameen Bakhtiary
'       - *Last modified:*  2012-01-12
'===================================================================================================
Private Sub emailSubmitter_Click()
    
    submitterName = Split(submitter, " ")
    MsgBox submitterName1 & vbNewLine & submittername2
    
'    Dim olApp As Object
'    Dim olMsg As Object
'
'    Set olApp = GetObject(, "Outlook.Application")
'    If olApp Is Nothing Then
'        Set olApp = CreateObject("Outlook.Application")
'    End If
'
'    Const olMailItem = 0
'    Set olMsg = olApp.CreateItem(olMailItem)
'    With olMsg
'        .To = {STRING:To}
'        .Subject = {STRING:Subject}
'        .Body = {STRING:Body}
    '     .Attachments.Add AttachmentArray(j)
    
    'Decide what to do with the mailitem
    '    .Send
    '    .Display
    '    .Save   'New messages always save in Drafts, regardless of what folder you use for the Items.Add method.
    End With
    
    
        
        
End Sub






Private Sub resetME_Click()
    Call UserForm_Initialize
End Sub



'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: changeMade
'       Make note when any change is made to the GUI and mark it as NOT SAVED. Notify the user
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
    saved = False
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

Private Sub openRTA_Click()
    Call openCWIpage
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
'
'   -
'
' Parameters:
'   -
'
' Last Modified: 2012-01-14
' ___________________________________________________________________________________________________
' ===================================================================================================
Sub openCWIpage(Optional view As String = "rta")
    progP = myPath & "\Include\CMDline_Functions.exe"
    formattedRTAnum = Right(rtaNUm.Caption, 6)
    
    'Can't find CMDline_Functions
    '===============================
    If Dir(progP) = "" Then Call MsgBox("Uh oh... An important file couldn't be found:  CMDline_Functions.exe" & vbCrLf & "" & vbCrLf & "Without it, you cannot open RTAs directly in CWI. This file should be located" & Chr(10) & "in the Include folder within this worksheets directory. " & vbCrLf & vbCrLf & "Running the installer file should solve this issue", vbCritical Or vbSystemModal, "--=[ WD RTA Sheet ]=--")
    
    
    'In edit mode... check for changes & save b4 closing GUI
    '=========================================================
    If Application.Range("sheetviewmode") = "Edit" Then
        If saved = False Then Call uploadToCWI_Click
    End If
    
    'Close the GUI and open IE
    '============================
    Unload Me
    Call Shell("""" & progP & """ " & formattedRTAnum & " " & view, vbNormalFocus)
    Exit Sub
End Sub





'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: closeME_Click
'       Check that all changes have been saved and close the GUI
'
'====================================================================================================
Private Sub closeME_Click()
    If saved = False And Application.Range("sheetviewmode") <> "PMT" Then
        ans = MsgBox("" & vbCrLf & "You made changes to this RTA that have not been saved." & vbCrLf & "" & vbCrLf & "DISCARD CHANGES?", _
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
    
    ' Formatted RTA Number R00000XXXXXX
    '==================================
    tmp = "R00000" & Strings.Right(rtaNUm.Caption, 6)
    
    
    ' Find the first open cell on RTAimport sheet or find the same
    ' RTA already on the sheet to overwrite
    '===============================================================
    r = 1
    While Cells(r, 1) <> ""
        If Cells(r, 2) = tmp Then GoTo overwrite
        r = r + 1
    Wend
    
    
overwrite:
    Range("a" & r) = "Rta"
    Range("b" & r) = tmp
    
    
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




'___________________________________________________________________________________________________
'***************************************************************************************************
'   Sub: wmComment_Click
'       Inserts the formatted weekly meeting comment start (YYYY-MM-DD, WM: ) at the bottom of the
'       description
'
'   Group: About
'       - *Written by:* Rameen Bakhtiary
'       - *Last modified:*  2012-01-10
'___________________________________________________________________________________________________
'***************************************************************************************************
Private Sub wmComment_Click()
    
    If Application.Range("sheetviewmode") = "PMT" Then Call MsgBox( _
        "You must be in Edit mode to insert a weekly meeting comment!", _
        vbExclamation Or vbSystemModal, "Insert Weekly Meeting Comments"): Exit Sub

    insertTxt = Chr(10) & Chr(10) & Format(Now(), "yyyy-MM-dd") & ", WM: "
    desc = desc & insertTxt
    desc.SelStart = 5000
    
    
End Sub










