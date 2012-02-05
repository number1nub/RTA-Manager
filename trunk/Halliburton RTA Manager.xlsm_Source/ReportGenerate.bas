Attribute VB_Name = "ReportGenerate"
'_______________________________________________________________________________________________________________
'===============================================================================================================
'   Title: Create PDF RTA Priorities Report
'
'   Group: Overview
'
'   Group: About
'       General script information
'       - *Written by:* Rameen Bakhtiary
'       - *Created on:* 2012-01-10
'_______________________________________________________________________________________________________________
'===============================================================================================================



    '____________________________________________________________________________________________________
    '====================================================================================================
    ' Sub:          moveToXML
    '
    ' Description:
    '               Generate and save a report of the currently visible RTAs which have priorities
    '               assigned.
    ' Parameters:
    '               None.
    ' Remarks:
    '               Launched by clicking the PDF report button on main sheet.
    ' Prerequisite:
    '               Must be in PMT mode, and must have the table sorted by a lab office.
    '
    '====================================================================================================
    Sub moveToXML(Optional reportType = "reportHide")
        Range("Main[[#Headers],[ ]]").Select    ' Select header cell of priority column
        
        
        
        ' Make sure that a lab office filter is set and that sheet is in PMT mode; else abort
        '======================================================================================
        If ActiveSheet.AutoFilter Is Nothing Or _
            Application.Range("sheetviewmode") <> "PMT" Then: Call _
            MsgBox("To generate a PDF report of RTA priorities you must:" & vbCrLf _
            & "" & vbCrLf & "       - Be in PMT mode" & vbCrLf & _
            "       - Select / filter by one of the lab-offices", vbExclamation Or _
            vbSystemModal, "Generate RTA Priority Report"): End
        Application.ScreenUpdating = False
        
        
        
        'Show columns to be in report & sort by priority
        '======================================================
        Cells.EntireColumn.Hidden = False
        Application.Range(reportType).EntireColumn.Hidden = True
        Call sortField(" ", "A")
        
        
        
        
        
        ' Find row number of last visible cell containing a priority
        '=============================================
       ActiveSheet.AutoFilter.Range.Offset(1, _
           0).Resize(ActiveSheet.AutoFilter.Range.Rows.Count - 1, _
           1).SpecialCells(xlCellTypeVisible).Select
       firstrw = ActiveCell.Row
        For Each rw In Selection
            If rw = "" Then: lastrw = Previous: Exit For
            Previous = rw.Row
            lastrw = rw.Row
        Next rw
        
        
    
        'Get Lab Office name
        '=================
        Select Case Cells(lastrw, getCol("Lab Office"))
        Case "WD3"
            loname = "Permanent Monitoring"
        Case "WD1"
            loname = "Flow Control"
        Case "WD4"
            loname = "Flow Control"
        Case "WD2"
            loname = "Digital Infrastructure"
        Case "WD5"
            loname = "Software"
        End Select
        
        
        '==== Show/Hide desired report columns ================
        ActiveSheet.Range("Main[Assigned To]").EntireColumn.Hidden = False
        ActiveSheet.Range("Main[Standard Production Lead Time]").EntireColumn.Hidden _
            = True
        ActiveSheet.Range("Main[Remaining Production Lead Time]").EntireColumn.Hidden _
            = True
        
        '==== Copy/paste to XML table =========================
        ActiveSheet.Range("A" & firstrw & ":R" & lastrw).Select
        Selection.Copy
        Sheets("Report").Visible = True
        Sheets("Report").Select
        Range("Table16").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        
        ' Header / File Name
        '==================
        fname = Format(Now(), "m/d/yyyy") & " " & loname & _
            " RTA Priorities List"
        fFullPath = myPath & "\Generated Reports\" & Format(Now(), _
            "m-d-yyyy") & " " & loname & " Sustaining Priorities.pdf"
        Range("a1") = fname
        
        ' Make sure that the Generated Reports folder exists; else create it
        '===================================================
        If Dir(myPath & "\Generated Reports", vbDirectory) = "" Then _
            MkDir (myPath & "\Generated Reports")
        
        'Save to PDF
        '=============
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fFullPath, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=True
        
        ' Clear report table & return to main sheet
        '=================================
        Rows(3).Select
        ActiveSheet.Range(Selection, Selection.End(xlDown)).Select
        Selection.Delete
        Sheets("RTA Manager").Select
        Sheets("Report").Visible = False
        Application.CutCopyMode = False
        Range("A1").Select
        ActiveSheet.Range("Main[Assigned To]").EntireColumn.Hidden = True
        ActiveSheet.Range("Main[Standard Production Lead Time]").EntireColumn.Hidden _
            = False
        ActiveSheet.Range("Main[Remaining Production Lead Time]").EntireColumn.Hidden _
            = False
        
        'Prompt that PDF report was saved
        '=============================
        promptTxt = _
            "PDF priority list was saved in the Generated Reports folder."
        Call CMDline("/popup", promptTxt, loname & " Priority Report")
        Application.ScreenUpdating = True
    End Sub

