Attribute VB_Name = "PrioritizeR"
'############################################################################################
'
'                                                               PRIORITY HANDLING FUNCTIONS
'
'           Written by:     Rameen Bakhtiary
'           Date:               7/25/2011
'
'           Description:    Functions used to manage the RTAs assigned priorities within one lab office
'
'############################################################################################


' PUBLIC  VARIABLES
'=====================
Public setPriority As Variant


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                               prioritize_ReOrder
'
'           Description:        Re-number all showing priorities starting at 1; from the lowest to highest

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub prioritize_ReOrder()
    '==================================================
    '                               Verify that all req'd conditions are met
    '==================================================
    ' Ensure that sheet is in EDIT mode; else abort.
    If Application.Range("sheetviewmode") = "PMT" Then Call MsgBox("You must be in EDIT mode in order to use this feature", vbCritical, "             ---===[ Prioritizer ]===---"): End
    
    ' PROMPT user to make sure that filters are set so that all prioritized RTAs for the given lab office are showing b4 re-numbering
    ans = MsgBox("MAKE SURE that all RTAs that have priorities assigned for this lab office are being shown in the current table." & vbCrLf & "" & vbCrLf & "Continue with re-numbering the currently displayed RTAs?", vbYesNo Or vbExclamation Or vbSystemModal Or vbDefaultButton1, "    ---====[ THE PRIORITIZER ]====---")
    If ans = vbNo Then Exit Sub
    
    ' Make sure that no RTAs have been changed & moved to the RTAimport sheet for loading to CWI; prompt and abort if there are
    If Sheet3.Range("A2") <> "" Then Call MsgBox("You have other in-process RTA changes that haven't been saved to CWI..." & vbCrLf & "" & vbCrLf & _
        "Before using the prioritizer to re-number RTA priorities, please 'Finish' current RTA changes, and load" & vbCrLf & "them to CWI." & vbCrLf & "" & vbCrLf & _
        "After loading your changes to CWI, please give the system enough time to update/refresh changes so that" & vbCrLf & _
        "when refreshed, this sheet reflects the changes made in CWI; then re-run the Prioritizer.", vbCritical Or vbSystemModal Or vbDefaultButton1, "    ---====[ THE PRIORITIZER ]====---"): Exit Sub
    
    '==================================================
    '                                                  Begin Prioritizer Code
    '==================================================
    Application.ScreenUpdating = False
    Call sortField(" ")     'Ensure that table is sorted by priority
    pRows = LastFullRow()       'Array of all currently VISIBLE row nums that have priorities
    numRows = total_Rows(pRows)     'Number of prioritized rows
    priorityCol = getCol(" "): commentcol = getCol("Comments")
    newnum = 1
    '=== Renumber the priorities ====
    For Each rw In pRows
        If Cells(rw, priorityCol) <> newnum Then
            curPriority = Cells(rw, priorityCol)
            Cells(rw, commentcol) = Strings.Replace(Cells(rw, commentcol), curPriority & ":", newnum & ":")
            '=== Copy to RTAload sheet ===
            writeToImport (rw)
        End If
        newnum = newnum + 1
        curPriority = ""
    Next
    '=== Save load sheet ====
    Call saveImportSheet
    '=== Clear CWI Import Sheet ====
    Sheet6.Visible = xlSheetVisible
    Sheet6.Select
    Sheet6.Range("a2").Select
    Sheet6.Rows("2:2").Select
    Sheet6.Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheet6.Range("a2").Select
    Sheets("RTA Manager").Select
    Sheet6.Visible = xlSheetHidden
    Application.Range("sheetviewmode") = "PMT"
    Cells.EntireColumn.Hidden = False
    Application.GoTo ("pmthide")
    Selection.EntireColumn.Hidden = True
    Sheet1.sheetView.Caption = "SHEET MODE: PMT"
    Sheet1.clearSort.Caption = "Reset"
    Sheet1.PM.Visible = True
    Sheet1.fc.Visible = True
    Sheet1.di.Visible = True
    Sheet1.soft.Visible = True
    Sheet1.pmSht.Visible = True
    Sheet1.fcSht.Visible = True
    Sheet1.diSht.Visible = True
    Sheet1.softSht.Visible = True
    selectTop (6)
    Sheet1.Range("a1").Select
    Call MsgBox("All Done!" & vbCrLf & "" & vbCrLf & "Press OK to continue and load the changes to CWI." & vbNewLine & vbNewLine & "Remember that changes take some time to update from CWI and reflect in the sheet.", vbInformation, "    ---====[ THE PRIORITIZER ]====---")
    Call Shell("""" & ThisWorkbook.Path & "\Include\CMDline_Functions.exe"" /Load", vbNormalFocus)
    Application.ScreenUpdating = True
End Sub





'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 writeToImport
'
'           Description:        Writes a row's info to the RTAimport sheet

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub writeToImport(rwNum As Integer)
    r = 1
    While Sheet6.Cells(r, 1) <> ""
        r = r + 1
    Wend
    Sheet6.Range("a" & r) = "Rta"
    fullRtaNum = "R00000" & Sheet1.Cells(rwNum, getCol("RTA"))
    Sheet6.Range("b" & r) = fullRtaNum
    Sheet6.Range("c" & r) = Sheet1.Cells(rwNum, getCol("Comments"))
End Sub



'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 saveImportSheet
'
'           Description:        Save the import sheet to a new workbook in My Documents folder named rtaLoad.xlsx

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub saveImportSheet()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Sheet6.Visible = xlSheetVisible
    Sheet6.Select
    Sheet6.Copy
    un = UserNameWindows
    ChDir "C:\documents and settings\" & un & "\my documents\"
    ActiveWorkbook.SaveAs Filename:="C:\documents and settings\" & un & "\my documents\rtaLoad.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Sheet6.Visible = xlSheetHidden
End Sub





'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 prioritize_Single
'
'           Description:        Allows user to set or change the priority of an RTA by double-clicking its priority cell

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub prioritize_Single(cRow, curPriority)
    Application.ScreenUpdating = False
    prioritize.Show
    
    ' User pressed cancel or was not in EDIT mode; abort
    '==========================================
    If setPriority = "Cancel" Then Exit Sub
    If Application.Range("sheetviewmode") = "PMT" Then Call MsgBox("Oops..." & vbCrLf & "" & vbCrLf & _
        "You must be in ""EDIT"" mode in order to make changes to RTA priorities.", vbCritical Or vbSystemModal, "        ---==[ The Prioritizer ]==---"): End
    
    'Priority entered is same as previous; abort
    '==================================
    If setPriority = curPriority Then Exit Sub
    
    ' Current priority is blank -- append new priority to front of comment
    '=====================================================
    If curPriority = "" And setPriority <> "" Then
        Cells(cRow, getCol("Comments")) = setPriority & ": " & Cells(cRow, getCol("Comments"))
    
    ' Current priority exists
    '==================
    Else
        'Add colon if new priority isn't blank; This allows entering a blank priority to erase the previous w/its colon
        If setPriority <> "" Then setPriority = setPriority & ":"
        
        'Replace previous priority with entry; erases priorty if entry was blank
        '=====================================================
        Cells(cRow, getCol("Comments")) = Replace(Cells(cRow, getCol("Comments")), curPriority & ":", setPriority)
    End If
    
    ' Write to RTAimport sheet for uploading to CWI
    '======================================
    fullRtaNum = "R00000" & Sheet1.Cells(ActiveCell.Row, getCol("RTA"))    'RTA number with R00000 added to front
    
    'Get row number of first empty cell on load sheet OR current RTA already on the sheet
    r = 1
    While Sheets("RTAimport").Cells(r, 1) <> ""
        If Sheet3.Cells(r, 2) = fullRtaNum Then GoTo overwrite
        r = r + 1
    Wend
overwrite:
        Sheet3.Range("a" & r) = "Rta"
        Sheet3.Range("b" & r) = fullRtaNum
        Sheet3.Range("c" & r) = Sheet1.Cells(ActiveCell.Row, getCol("Description"))
        Sheet3.Range("d" & r) = Sheet1.Cells(ActiveCell.Row, getCol("Comments"))
        
        ' Convert simple RTA class to CWI format
        '=================================
        Select Case Sheet1.Cells(ActiveCell.Row, getCol("Class"))
        Case "A"
            fullc = "A=Minimal Processing Time"
        Case "B"
            fullc = "B=Medium Processing Time"
        Case "C"
            fullc = "C=Technology Negotiated Processing Time"
        Case "D"
            fullc = "D=Technology Development Engineering"
        End Select
        
        Sheet3.Range("e" & r) = fullc
        Sheet3.Range("f" & r) = Sheet1.Cells(ActiveCell.Row, getCol("Assigned To"))
        Sheet3.Range("g" & r) = Sheet1.Cells(ActiveCell.Row, getCol("Current Status"))
        Sheet3.Range("h" & r) = Sheet1.Cells(ActiveCell.Row, getCol("Revised Due Date"))
        
        ' Save RTAimport sheet to file in user's My Documents folder named rtaLoad.xlsx
        '==============================================================
        Application.DisplayAlerts = False
        Sheets("RTAimport").Visible = xlSheetVisible
        Sheets("RTAimport").Select
        Sheets("RTAimport").Copy
        un = UserNameWindows        'Get Windows username
        ChDir "C:\documents and settings\" & un & "\my documents\"
        ActiveWorkbook.SaveAs Filename:="C:\documents and settings\" & un & "\my documents\rtaLoad.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWorkbook.Close        'Close the newly created file
        Application.DisplayAlerts = True
        Sheet3.Visible = xlSheetHidden
    Application.ScreenUpdating = True
End Sub

























