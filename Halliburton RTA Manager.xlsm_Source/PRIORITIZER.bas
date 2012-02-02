Attribute VB_Name = "PRIORITIZER"
' ___________________________________________________________________________________________________
' ***************************************************************************************************
' Module: Prioritizer
'
'   This module contains the functions and routines used in "The Prioritizer."
'
'   When run, the prioritizer will-
'       o   Sort a given lab office's RTAs by priority
'       o   Determine the number of RTAs with priorities assigned, and create an
'           array with the row number of each
'       o   Loop through each RTA, renumbering each (if it needs it) starting from 1.
'       o   An rtaLoad.xlsx sheet is created and saved into user's My Documents folder
'           containing on the Rta Name column and Comments column (on a sheet named
'           priorityLoad).
'
'
' About:
'   o Written by:    Rameen Bakhtiary
'   o Last Modified: 2012-01-16
' ___________________________________________________________________________________________________
' ***************************************************************************************************


Public setPriority As Variant




'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: prioritize_ReOrder
'       Run "The Prioritizer" to re-number all showing priorities starting from 1
'
'   Group: Remarks
'       Run by clicking "Prioritizer" button on main sheet
'
'====================================================================================================
Sub prioritize_ReOrder()
   
    '==================================================
    '         Verify that all req'd conditions are met
    '==================================================
        
        '_________________________________________________
        '   Ensure that sheet is in EDIT mode; else abort.
        '
        If Application.Range("sheetviewmode") = "PMT" Then Call _
            MsgBox("You must be in EDIT mode in order to use this feature", _
            vbCritical, "             ---===[ Prioritizer ]===---"): End
        
        
        '______________________________________________________________________________
        '   PROMPT  USER TO MAKE SURE THAT FILTERS ARE SET SO THAT ALL PRIORITIZED RTAS
        '   FOR THE GIVEN LAB OFFICE ARE SHOWING B4 RE-NUMBERING
        '
        ans = _
            MsgBox("MAKE SURE that all RTAs that have priorities assigned for this lab office are being shown in the current table" _
            & vbCrLf & "and that the table is sorted (ascending) by priority." & _
            vbCrLf & "" & vbCrLf & _
            "Continue with re-numbering the currently displayed RTAs?", vbYesNo Or _
            vbExclamation Or vbSystemModal Or vbDefaultButton1, _
            "    ---====[ THE PRIORITIZER ]====---")
        If ans = vbNo Then Exit Sub
        
        '______________________________________________________________________________________________
        '   MAKE SURE THAT NO RTAS HAVE BEEN CHANGED & MOVED TO THE RTAIMPORT SHEET FOR LOADING TO CWI;
        '   PROMPT AND ABORT IF THERE ARE
        '
        If Sheet3.Range("A2") <> "" Then Call _
            MsgBox("You have other in-process RTA changes that haven't been saved to CWI..." _
            & vbCrLf & "" & vbCrLf & _
            "Before using the prioritizer to re-number RTA priorities, please 'Finish' current RTA changes, and load" _
            & vbCrLf & "them to CWI." & vbCrLf & "" & vbCrLf & _
            "After loading your changes to CWI, please give the system enough time to update/refresh changes so that" _
            & vbCrLf & _
            "when refreshed, this sheet reflects the changes made in CWI; then re-run the Prioritizer.", _
            vbCritical Or vbSystemModal Or vbDefaultButton1, _
            "    ---====[ THE PRIORITIZER ]====---"): Exit Sub
        
    Application.ScreenUpdating = False
    
    
    
    ' ===================================================
    ' PREPARATION
    ' ===================================================
        
        '____________________________
        '   DISPLAY THE SPLASH SCREEN
        '
        Call splash("The RTA list is being re-numbered. " & vbNewLine & "Please wait...")
        
        
        '_________________________
        '   SORT TABLE BY PRIORITY
        '
        Call sortField(" ")
        
        
        '__________________________________________________________________
        '   GET AN ARRAY OF ALL CURRENTLY VISIBLE ROWS THAT ARE PRIORITIZED
        pRows = LastFullRow()
        
        
        '___________________________________
        '   NUMBER OF ROWS TO BE PRIORITIZED
        '
        numRows = total_Rows(pRows)
        
        '__________________________________________________
        '   GET COLUMN NUMBER OF PRIORITY & COMMENTS COLUMN
        '
        priorityCol = getCol(" "): commentcol = getCol("Comments")
        
    
    

    ' ===================================================
    ' PERFORM  RE-NUMBERING
    ' ===================================================
    newnum = 1
    For Each rw In pRows
        
        '____________________________________________________
        '   CHANGE THE PRIORITY NUMBER UNLESS ALREADY CORRECT
        If Cells(rw, priorityCol) <> newnum Then
            curPriority = Cells(rw, priorityCol)
            Cells(rw, commentcol) = Strings.Replace(Cells(rw, commentcol), curPriority & ":", newnum & ":")
            
            '_________________________
            '   WRITE TO RTALOAD SHEET
            '
            writeToImport (rw)
        End If
        
        newnum = newnum + 1
        curPriority = ""
    Next
        
    
    '===================================================
    '       SAVE RTALOAD SHEET TO MY DOCUMENTS
    '===================================================
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
    Sheet1.ClearSort.Caption = "Reset"
    Sheet1.PM.Visible = True
    Sheet1.fc.Visible = True
    Sheet1.di.Visible = True
    Sheet1.soft.Visible = True
    Sheet1.pmSht.Visible = True
    Sheet1.fcSht.Visible = True
    Sheet1.diSht.Visible = True
    Sheet1.softSht.Visible = True
    
            
    'Clear the filter and exit Edit mode
    '=====================================
    Range("Main[[#Headers],[RTA]]").Select
    Selection.AutoFilter
    ActiveSheet.gotoDept.Visible = False
    
    
    'Refresh the table
    '===================
    Selection.ListObject.QueryTable.refresh
    Call selectTop
    Sheet1.Range("a1").Select
    
    
    'Turn off splash
    '=================
    splash
        
    
    'Warn about refreshing data
    '==============================
    Call _
        MsgBox("Prioritization complete.  rtaLoad.xlsx is saved in My Documents for loading into CWI." _
        & vbCrLf & "" & vbCrLf & "    ******** !!!!!   IMPORTANT   !!!!!! *******" _
        & vbCrLf & "" & vbCrLf & "CHANGES LOADED THROUGH ""MODIFY OBJECTS FROM EXCEL"" TAKE A FEW MINUTES TO UPDATE AND SHOW UP IN THE RTA SHEET.  " & vbCrLf & "" & vbCrLf & "DONT MAKE ANY CHANGES USING THE RTA SHEET UNTIL YOU SEE YOUR CHANGES REFLECTED (PERIODICALLY REFRESH THE DATA)" & vbCrLf & "" & vbCrLf & "    ********************************* ", _
        vbInformation Or vbSystemModal, "Prioritizer:  ALL DONE!")

    'Display a notification of completion
    '=====================================
    Call CMDline_Func("/popup", "          Load  'My Documents\rtaLoad.xlsx'  into CWI to apply the changes" & vbNewLine & vbNewLine, "          Prioritization complete....")
    
    'Open a CWI page
    '=================
    Call CMDline_Func("/Load", vbNormalFocus)
            
    Application.ScreenUpdating = True
End Sub









'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: writeToImport
'
'       Write new RTA information to the PriorityLoad sheet. Used when running the
'       prioritizer.
'
'   Parameters:
'       rwNum   -   Row number of the current RTA on main sheet. Used to get RTA info
'
'====================================================================================================
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








'____________________________________________________________________________________________________
'====================================================================================================
'   Sub: saveImportSheet
'
'       Save the import sheet to a new workbook in My Documents folder named rtaLoad.xlsx
'
'====================================================================================================
Sub saveImportSheet()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '_____________________
    '   COPY RTALOAD SHEET
    '
    Sheet6.Visible = xlSheetVisible
    Sheet6.Select
    Sheet6.Copy
    
    '______________________________
    '   SAVE IN MY DOCUMENTS FOLDER
    '
    ChDir MyDocs
    ActiveWorkbook.SaveAs Filename:=MyDocs & "rtaLoad.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    '________________________________________________
    '   CLOSE THE NEWLY CREATED RTALOAD.XLSX WORKBOOK
    '
    ActiveWorkbook.Close
    
    
    '_______________________________________________________
    '   SAVE A COPY OF THE RTALOAD SHEET TO THE PUBLIC DRIVE
    '   FOR USE AS A BACKUP
    '
    Sheet6.Select
    Sheet6.Copy

    tStamp = Format(Now(), "yyyy-m-d hhmm")
    ChDir BACKUP_DIR
    fname = BACKUP_DIR & tStamp & " (" & un & ").xlsx"
    
    ActiveWorkbook.SaveAs Filename:=fname, FileFormat:=xlOpenXMLWorkbook
    
    
    '________________________________________________
    '   CLOSE THE NEWLY CREATED RTALOAD.XLSX WORKBOOK
    '
    ActiveWorkbook.Close
            
    Application.DisplayAlerts = True
    Sheet6.Visible = xlSheetHidden
End Sub






'__________________________________________________________________________________________________________
'==========================================================================================================
' Function:     LastFullRow
'
'               Returns an array of row numbers that are visible in the current table and whose value in
'               the given column are not blank.
' Parameters:
'               [colHeader] -   (Optional) Column header title of column to whose contents musn't be blank
'                               Defualt is the priority column (header: " ")
'
'==========================================================================================================
Function LastFullRow(Optional colHeader As String = " ") As Variant
    Application.ScreenUpdating = False
    '==== Select all visible cells in column ============
    Call selectVisible(colHeader)
    '==== Create Comma Separated list of rows =======
    For Each rw In Selection
        If rw = "" Then Exit For
        If rowlist = "" Then
            rowlist = rw.Row
        Else
            rowlist = rowlist & ", " & rw.Row
        End If
    Next
    '=== Return =============
    LastFullRow = Split(rowlist, ", ")
    Range("a1").Select
    Application.ScreenUpdating = True
End Function




'____________________________________________________________________________________________________
'====================================================================================================
' Function:     selectVisible
'
'               Selects all VISIBLE cells in a given column that aren't empty
' Parameters:
'               -
'
'====================================================================================================
Function selectVisible(colName As String)
    cCol = getCol(colName)
    Range("Main[[#Headers],[" & colName & "]]").Select
    '==== No filter; All rows showing ================
    If ActiveSheet.AutoFilter Is Nothing Then
        Selection.Offset(1, 0).Select
        i = Selection.Row: Top = Selection.Row
        While Cells(i, cCol).Offset(1, 0) <> ""
            ActiveSheet.Range(Selection, Selection.Offset(1, 0)).Select
            i = i + 1
        Wend
    '==== Filter on; Only select non-hidden rows =========
    Else
        ActiveSheet.AutoFilter.Range.Offset(1, 0).Resize(ActiveSheet.AutoFilter.Range.Rows.Count - 1, 1).SpecialCells(xlCellTypeVisible).Select
    End If
End Function






'____________________________________________________________________________________________________
'====================================================================================================
' Function:     last_Row
'
'               Returns the number of the last visible row in the current table that has a priority
'====================================================================================================
Function last_Row() As Integer
    last_Row = LastFullRow(total_Rows() - 1)
End Function





'____________________________________________________________________________________________________
'====================================================================================================
' Function:     total_Rows
'
'               Returns the number of items in a given array.
' Parameters:
'               [arry]  -   (Optional) Array whose length to return.
'                           Default array used is obtained by calling LastFullRow() function.
'
'====================================================================================================
Function total_Rows(Optional arry As Variant = "") As Integer
    total_Rows = 0
    If IsEmpty(arry) Then arry = LastFullRow()
    For Each itm In arry
        total_Rows = total_Rows + 1
    Next
End Function
















