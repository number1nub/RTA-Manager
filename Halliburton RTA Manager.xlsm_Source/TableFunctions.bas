Attribute VB_Name = "TableFunctions"
'___________________________________________________________________________________________________
'***************************************************************************************************
'   Title: FILTER_AND_SORT
'---------------------------------------------------------------------------------------------------
'   Group: Overview
'       General overview of script features, functions & implementation
'
'       Functions that handle filtering/sorting the RTA table. This includes getting the custom
'       settings for each lab office.
'
'   Group: About
'       General script information
'       - *Written by:* Rameen Bakhtiary
'       - *Created on:* 2012-01-10
'___________________________________________________________________________________________________
'***************************************************************************************************




    
    '____________________________________________________________________________________________________
    '====================================================================================================
    ' Function:     getSettings
    '
    ' Written by:   Rameen Bakhtiary
    ' Created on:   10/24/2011
    ' Description:
    '               Return an array of values in a given named range.
    '               Used to get the custom filter setting values for a lab office for table sorting
    ' Parameters:
    '               rgName - Name of named-range in current sheet
    '
    '====================================================================================================
    Function getSettings(rgname As Variant) As Variant
        'Each customizable value can have up to 10 filter items
        Dim farray(1 To 10) As String
        Set newrg = Application.Range(rgname)
        
        
        i = 1
        For Each itm In newrg
            If itm <> "" Then
                farray(i) = itm
                i = i + 1
            End If
        Next
        getSettings = farray
    End Function
    
    
      
      


    '____________________________________________________________________________________________________
    '====================================================================================================
    ' Sub:          sortBy  (Main table sort routine)
    '
    ' Written by:   Rameen Bakhtiary
    ' Description:
    '               Applies filters to main RTA table based on custom lab office settings when one of
    '               the lab office buttons is pressed
    ' Parameters:
    '               prefix - String prefix of lab office to filter by
    ' Remarks:
    '               Relies on getSettings() and getCol() functions
    '
    '====================================================================================================
    Sub sortBy(prefix As String)
        If Application.Range("inproc") = 1 Then Exit Sub
        Application.ScreenUpdating = False
        Application.Range("inProc") = 1
        
        
        '=============================================================
        '        Remove the filter drop downs from headers
        '=============================================================
        For Each c In Range("Main[#Headers]")
            ActiveSheet.ListObjects("Main").Range.AutoFilter Field:=getCol(c), visibledropdown:=False
        Next
        
        
        
        
        '_______________________________  APPLY FILTER SETTINGS TO TABLE ___________________________________
        '%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%&%
                    
        '_________________
        '       RTA  STATE
        '
        ActiveSheet.ListObjects("Main").Range.AutoFilter Field:=getCol("State"), _
            Criteria1:=getSettings(prefix & "State"), Operator:=xlFilterValues
        
        '_________________
        '       LAB OFFICE
        '
        ActiveSheet.ListObjects("Main").Range.AutoFilter Field:=getCol("Lab Office"), _
            Criteria1:=getSettings(prefix & "Lab"), Operator:=xlFilterValues
        
        '_______________
        '       RTA TYPE
        '
        ActiveSheet.ListObjects("Main").Range.AutoFilter Field:=getCol("Type"), _
            Criteria1:=getSettings(prefix & "Type"), Operator:=xlFilterValues
        
        '_______________
        '       RTA CODE
        '
        ActiveSheet.ListObjects("Main").Range.AutoFilter Field:=getCol("Code"), _
            Criteria1:=getSettings(prefix & "Code"), Operator:=xlFilterValues
        '___________________________________________________________________________________________________
        
    
    
        If Application.Range("sheetviewmode") = "PMT" Then ActiveSheet.gotoDept.Visible = True
        Application.Range("cfilt") = prefix
        Application.Range("inProc") = 0
        ActiveSheet.Range("a1").Select
        Application.ScreenUpdating = True
    End Sub




    

'____________________________________________________________________________________________________
'====================================================================================================
' Sub:      sortField
'
' Written by:   Rameen Bakhtiary
' Created on:   10/24/2011
' Description:
'               Sorts the table by a certain column given its header text. Toggles between ascending
'               and Descending sort order when the currently sorted column is passed again
' Parameters:
'               fieldTitle - Header text of column to sort
'               [sOrder] - (Optional) Either "A" or "D" for ascending/descending sort order
'                           Default is ascending ("A")
'
'====================================================================================================
Sub sortField(fieldTitle As String, Optional sOrder As String = "A")
    Application.ScreenUpdating = False
        
    'Clear current sort
    ActiveWorkbook.ActiveSheet.ListObjects("Main").Sort.SortFields.Clear
    
    '============================================================
    '   SORT ASCENDING - sOrder not specified, invalid or "A"
    '============================================================
    If sOrder <> "D" Then
        ActiveWorkbook.ActiveSheet.ListObjects("Main").Sort.SortFields.Add Key:=Range("Main[[#All],[" & fieldTitle & "]]"), _
         SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    '===================================================
    '   SORT DESCENDING - sOrder = "D"
    '===================================================
    Else
        ActiveWorkbook.ActiveSheet.ListObjects("Main").Sort.SortFields.Add Key:=Range("Main[[#All],[" & fieldTitle & "]]"), _
         SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    End If
    
    With ActiveWorkbook.ActiveSheet.ListObjects("Main").Sort
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
    End With
End Sub




    


'____________________________________________________________________________________________________
'====================================================================================================
' Sub:          selectTop
'
' Written by:   Rameen Bakhtiary
' Created on:   10/24/2011
' Description:
'               Selects the top-left-most visible cell in the table.
' Parameters:
'               [col]       -   Column number of "left-most" column to select. Must pass as integer
'                               Default is 3 (column "c")
'               [rowMax]    -   Max row number value to search to when looking for a non-hidden
'                               cell. Used to conserve time/memory in situations where all rows
'                               may have been hidden.
'
'====================================================================================================
Sub selectTop(Optional col = 6, Optional rowMax = 2000)
    
    'Get row number of top-most cell in table (when no rows are hidden)
    '=====================================================
    i = Range("Main[#Headers]").Row + 1
    
    'Find first visible row, starting at topmost cell
    '===================================
    While Rows(i & ":" & i).EntireRow.Hidden
        i = i + 1
        If i > rowMax Then
            Exit Sub
        End If
    Wend
    Cells(i, col).Select
End Sub

    







