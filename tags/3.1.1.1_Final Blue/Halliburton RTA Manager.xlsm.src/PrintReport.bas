Attribute VB_Name = "PrintReport"

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'                                                                   GENERATE  PDF  REPORT  FOR  A  LAB  OFFICE
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Sub moveToXML()
    Range("Main[[#Headers],[ ]]").Select    ' Select header cell of priority column
    
    ' Make sure that a lab office filter is set and that sheet is in PMT mode; else abort
    '==============================================================
    If ActiveSheet.AutoFilter Is Nothing Or Application.Range("sheetviewmode") <> "PMT" Then: Call MsgBox("To generate a PDF report of RTA priorities you must:" & vbCrLf & "" & vbCrLf _
        & "       - Be in PMT mode" & vbCrLf & "       - Select / filter by one of the lab-offices", vbExclamation Or vbSystemModal, "Generate RTA Priority Report"): End
    Application.ScreenUpdating = False
    
    ' Find row number of last visible cell containing a priority
    '=============================================
   ActiveSheet.AutoFilter.Range.Offset(1, 0).Resize(ActiveSheet.AutoFilter.Range.Rows.Count - 1, 1).SpecialCells(xlCellTypeVisible).Select
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
    ActiveSheet.Range("Main[Standard Production Lead Time]").EntireColumn.Hidden = True
    ActiveSheet.Range("Main[Remaining Production Lead Time]").EntireColumn.Hidden = True
    
    '==== Copy/paste to XML table =========================
    ActiveSheet.Range("A" & firstrw & ":R" & lastrw).Select
    Selection.Copy
    Sheets("Report").Visible = True
    Sheets("Report").Select
    Range("Table16").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Header / File Name
    '==================
    fName = Format(Now(), "m/d/yyyy") & " " & loname & " RTA Priorities List"
    fFullPath = ThisWorkbook.Path & "\Generated Reports\" & Format(Now(), "m-d-yyyy") & " " & loname & " Sustaining Priorities.pdf"
    Range("a1") = fName
    
    ' Make sure that the Generated Reports folder exists; else create it
    '===================================================
    If Dir(ThisWorkbook.Path & "\Generated Reports", vbDirectory) = "" Then MkDir (ThisWorkbook.Path & "\Generated Reports")
    
    'Save to PDF
    '=============
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fFullPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
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
    ActiveSheet.Range("Main[Standard Production Lead Time]").EntireColumn.Hidden = False
    ActiveSheet.Range("Main[Remaining Production Lead Time]").EntireColumn.Hidden = False
    
    'Prompt that PDF report was saved
    '=============================
    promptTxt = "PDF priority list was saved in the Generated Reports folder."
    Call notify(promptTxt, loname & " Priority Report")
    Application.ScreenUpdating = True
End Sub





