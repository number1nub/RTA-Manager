Attribute VB_Name = "Settings_Functions"
'==================================================
'                           Public Variables & Function declarations
'==================================================
Public OK As Boolean       'Used to check for valid password entry
Public sMyPassWord As String        'Password to unlock certain sheet features
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long



'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 Function:  UserNameWindows
'
'       Description:    Returns Windows username (as opposed to useing the registered name to MS Office)

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function UserNameWindows() As String
    Dim lngLen As Long
    Dim strBuffer As String
    Const dhcMaxUserName = 255
    strBuffer = Space(dhcMaxUserName)
    lngLen = dhcMaxUserName
    If CBool(GetUserName(strBuffer, lngLen)) Then
        UserNameWindows = Left$(strBuffer, lngLen - 1)
    Else
        UserNameWindows = ""
    End If
End Function



 

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 VerifyScreenResolution
'
'           Description:        Checks that screen resolution is at least set to minimum requirements (sent as parameters; defaul = 1280 x 1024)
'                                     otherwise, prompt user to change resolution and open system screen settings
'
'           Parameters:         xmin - Minimum required horizontal component of resolution
'                                      ymin - Minimum required vertical component of resolution

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'Sub VerifyScreenResolution(Optional xmin As Integer = 1280, Optional ymin As Integer = 1024)
'
'    Dim x  As Long
'    Dim y  As Long
'    Dim MyMessage As String
'    Dim MyResponse As VbMsgBoxResult
'    x = GetSystemMetrics(0)
'    y = GetSystemMetrics(1)
'    If x < xmin Or y < ymin Then
'        MyMessage = "Your current screen resolution is " & x & " X " & y & vbCrLf & "This program " & _
'        "was designed to run with a screen resolution of at least  1280 X 1024 and may not function properly " & _
'        "with your current settings. It is highly recommended that you change your resolution." & vbCrLf & "Would you like to open resolution settings now?"
'        MyResponse = MsgBox(MyMessage, vbExclamation + vbYesNo, "Screen Resolution")
'    End If
'    If MyResponse = vbYes Then
'        Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3")
'    End If
'
'End Sub



'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                               Function:  getSettings
'
'       Description:    Returns an array of all values in a named range of cells. Used to get the custom filter values set by user.
'
'       Parameters:     rgName - A Range of cells to convert to an array
'       Return:            Returns an array of string values

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function getSettings(rgname As Variant) As Variant
    Dim farray(1 To 10) As String   'Each customizable value can have up to 10 filter items (except for lead time critera)
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



'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 Function:  LastFullRow
'
'               Description:        Returns  an array of the row numbers that are visible in the current table and have contents (aren't empty)
'
'               Parameter:          Column header title of the column to consider. [Defualt is the priority column]

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
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

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 Function: selectVisible
'
'           Description:        Selects all VISIBLE cells in a given column that aren't empty

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
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


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 Function:  last_Row
'
'           Description:        Returns the row number of the last row with a visible, prioritized RTA
'

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function last_Row() As Integer
    last_Row = LastFullRow(total_Rows() - 1)
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 Function: total_Rows
'
'           Description:      Returns the number of items in an array.  Defaults to the length of priority col array LastFullRow()
'

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function total_Rows(Optional arry As Variant = "") As Integer
    total_Rows = 0
    If IsEmpty(arry) Then arry = LastFullRow()
    For Each itm In arry
        total_Rows = total_Rows + 1
    Next
End Function





'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                           sortBy()
'
'       Description:    Performs filtering on the table when the buttons are pressed. Relies on getSettings() [above]
                                    'and getCol() [below]

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub sortBy(prefix As String)
        If Application.Range("inproc") = 1 Then Exit Sub
        Application.ScreenUpdating = False
        Application.Range("inProc") = 1
        '........................................................ Apply Filters ........................................................
        'RTA  STATE
        ActiveSheet.ListObjects("Main").Range.AutoFilter Field:=getCol("State"), Criteria1:=getSettings(prefix & "State"), Operator:=xlFilterValues
        'LAB OFFICE
        ActiveSheet.ListObjects("Main").Range.AutoFilter Field:=getCol("Lab Office"), Criteria1:=getSettings(prefix & "Lab"), Operator:=xlFilterValues
        'RTA TYPE
        ActiveSheet.ListObjects("Main").Range.AutoFilter Field:=getCol("Type"), Criteria1:=getSettings(prefix & "Type"), Operator:=xlFilterValues
        'RTA CODE
        ActiveSheet.ListObjects("Main").Range.AutoFilter Field:=getCol("Code"), Criteria1:=getSettings(prefix & "Code"), Operator:=xlFilterValues
        
        '==== Hide drop downs for selected columns =========
        Dim ColsToHideDropDown As Variant
        ColsToHideDropDown = Array(" ", "RTA", "Description", "In Stock Date", "Standard Production Lead Time", "Remaining Production Lead Time", "Revised Due Date", "Class", "Tech Start Date", "Design", "Draft", "Check", "Approve", "Request Complete Date", "Revised Due Date")
        For Each itm In ColsToHideDropDown
'            MsgBox itm
            ActiveSheet.ListObjects("Main").Range.AutoFilter Field:=getCol(itm), visibledropdown:=False
        Next
        
        If Application.Range("sheetviewmode") = "PMT" Then ActiveSheet.gotoDept.Visible = True
        Application.Range("cfilt") = prefix
        Application.Range("inProc") = 0
        ActiveSheet.Range("a1").Select
        Application.ScreenUpdating = True
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                           getCol()
'
'       Description:    Returns the column number of a given column header in the Main RTA table

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function getCol(header As Variant) As Integer
    getCol = Range("Main[[#Headers],[" & header & "]]").Column
End Function


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 Priority
'
'       Description:    Looks for a number followed by a colon at start of string. Retrun the number as integer if found.
'                             Retruns 0 if no priority found.

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function priority(TheString As String) As Double
    If TheString = "" Then priority = 0: Exit Function
    tmp = InStr(1, TheString, ":", vbTextCompare)
    If tmp = 0 Then priority = 0: Exit Function
    If tmp < 6 Then priority = Left(TheString, tmp - 1)
End Function


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 AtAt
'
'       Description:    Checks for '@@' in string. Retrun true if found. False otherwise

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function AtAt(TheString As String) As Boolean
    tmp = Strings.InStr(1, TheString, "@@", vbTextCompare) ' Find "@@" in the string
    If tmp = 0 Then ' if there are no "@@"
        AtAt = 0
    ElseIf tmp <> 0 Then ' if there is an "@@"
        AtAt = 1
    End If
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 SPLT Function
'
'       Description:    Get Standard Production Lead Time from comments. Searches for SPLT=#### [days].
'                             -Is not case sensitive
'                             -Must not have space before (=) sign. (Ex: SPLT = 85 is WRONG, SPLT= 85 is OKAY)
'                             -Returns "N/A" by default if not found. Can be changed from formula call in sheet.

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function stdProdLT(comTxt As String, Optional valueIfMissing = 0) As Double
    comTxt = UCase(comTxt)
    Dim pos As Integer: pos = InStr(1, comTxt, "SPLT=", vbTextCompare)
    If pos = 0 Then stdProdLT = valueIfMissing: Exit Function                                'SPLT= Not found
    tmpStr = Trim(Right(comTxt, Len(comTxt) - pos - 4))
    tmp = InStr(1, tmpStr, " ", vbTextCompare)
        If tmp <> 0 Then tmpStr = Left(tmpStr, tmp)
    tmpStr = Trim(tmpStr)
    If Not IsNumeric(tmpStr) Or tmpStr = 0 Then
        tmpStr = 0
        stdProdLT = Format(tmpStr, "0.1")
    Else
        stdProdLT = Format(tmpStr, "general number"): Exit Function
    End If
End Function


    
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





'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                             Password Protect Function
'
'

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function GetPassWord(title As String, Optional tmpPW As String = "wdr74!")
    sMyPassWord = tmpPW
    protect.Caption = title
    protect.Show
    GetPassWord = OK
End Function



'=====================================================================================
'                                                                                   Notify
'
'           Description:        Disaply pop-up notification at bottom of screen
'           Parameters:
'                                   txt - Main notification text
'                                   [title] -   Notification title txt (default is blank)
'                                   [duration] - Time [seconds] to display pop-up (default is 5 sec)
'                                                    if duration = "f" pop-up will flash until clicked by user.
'                                   [run]   -   Action to run if user clicks notification before timeout

'=====================================================================================
Sub notify(txt As Variant, Optional title As String = "", Optional duration = "", Optional run = "")
    paramList = "/popUp " & """" & txt & """" & " " & """" & title & """" & " " & """" & duration & """" & " " & """" & run & """"
    Call Shell("""" & ThisWorkbook.Path & "\Include\CMDline_Functions.exe"" " & paramList, vbNormalFocus)
End Sub



'
''-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
''                                                                                 addToolTip
''
''           Description:        Adds a tooltip-like comment to the RTA cell containing general RTA info
'
''-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'Sub addToolTip()
'    Application.ScreenUpdating = False
'    cCol = getCol("RTA"): descCol = getCol("Description"): deptCol = getCol("Current Status"): assToCol = getCol("assigned to"): stateCol = getCol("State"): typeCol = getCol("Type")
'    On Error Resume Next
'
'
'    For Each rta In ActiveSheet.Range("main[RTA]")
'        curCom = "": curCom = rta.Comment.Text
'        If (curCom = "") And (Rows(rta.Row).EntireRow.Hidden = False) Then
'            '====== Simple Tooltip =======
'            If ActiveSheet.showTT = True Then
'                rta.AddComment ("[" & UCase(Cells(rta.Row, stateCol)) & "] " & Chr(10) & "--------------------------" & Chr(10) & Cells(rta.Row, typeCol) & Chr(10) & Cells(rta.Row, assToCol) & " (" & Replace(Replace(Cells(rta.Row, deptCol), "WD-", "", , , vbTextCompare), "Document Control", "TSG", , , vbTextCompare) & ")")
'            '===== Detailed Tooltip =======
'            Else
'                rta.AddComment ("[" & UCase(Cells(rta.Row, stateCol)) & "] " & rta.Text & " (" & Cells(rta.Row, typeCol) & ")  |  " & Cells(rta.Row, assToCol) & " (" & Cells(rta.Row, deptCol) & ")" & Chr(10) & "=================================" & Chr(10) & Cells(rta.Row, descCol))
'            End If
'
'            '===== Settings ==========
'            rta.Comment.Shape.Select True
'            With Selection
'                .AutoSize = False
'                .Shape.AutoShapeType = msoShapeRoundedRectangle
'                With .Font
'                    .FontStyle = "Arial"
'                    .Size = 10
'                    .Bold = True
'                    .ColorIndex = 2
'                    .TintAndShade = 0
'                End With
'                With .ShapeRange
'                    .Line.Visible = msoFalse
'                    .Height = 300
'                    .Width = 400
'                    .Fill.PresetGradient msoGradientHorizontal, 3, msoGradientLateSunset
'                End With
'            End With
'            rta.Comment.Visible = False
'        End If
'    Next
'    Range("a1").Select
'    Application.ScreenUpdating = True
'End Sub





'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                         selectTop()
'
'       Description:    Selects the top-most cell that is still showing when rows are hidden.
'                            Avoids being scrolled down on filter change
'
'       Parameter:
'                             [optional] col - Default Column to select is "C". Must pass as integer value of column
'                             [optional] rowMax - stops looking for a non-hidden cell after this row (usually for situations where all rows
'                                                            containing values were hidden. Change to higher number range requires it.

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub selectTop(topRow As Integer, Optional col = 3, Optional rowMax = 2000)
    
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



'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'                                                             TOGGLED  FULL-SCREEN MODE
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub CheckBox6887_Click()
        
    'Show a pop-up notification of current  full-screen mode
    '==========================================
    If Application.Range("fullScreenMode") = True Then
        fsMode = "ON"
        Application.DisplayFullScreen = True
    Else
        fsMode = "OFF"
        Application.DisplayFullScreen = False
    End If
    Call notify("", "Full-Screen mode: " & fsMode, 2)
End Sub




'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 nameFilterRanges
'
'       Description:    Creates named range for each row (length of 10) using the cell to the left
'                            of selected cell as the range name.
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Sub nameFilterRanges()
    i = 2
    Dim rgname As String
    While Range("c" & i) <> ""
        rgname = Range("c" & i)
        Range("d" & i & ":m" & i).Select
        tmp = ActiveWorkbook.Names.Add(rgname, Selection)
        i = i + 1
    Wend
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                  GetColLetter  [NUMBER =>> LETTER]
'
'       Description:    Returns a letter A - Z representing the input column number. Allows much simpler usage of "Range" command.

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function colNumToLetter(ColumnNumber As Integer) As String
    If ColumnNumber < 26 Then
        GetColumnLetter = Chr(ColumnNumber + 64)
    Else
        GetColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & Chr(((ColumnNumber - 1) Mod 26) + 65)
    End If
End Function



'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                        colLetterToNum  [ LETTER ==> NUMBER]
'
'       Description:    Returns the number of a column given the column letter

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function colLetterToNum(ByVal LtrIn As String) As Integer
    Dim TempChar As String
    Dim NumString  As String
    Dim TempNum As Integer
    Dim NumArray() As Integer
    Dim i As Integer
    Dim HighPower As Integer
    
    TempChar = ""
    TempNum = 0
    LtrIn = UCase(LtrIn)
    For i = 1 To Len(LtrIn)
        NumString = ""
        TempChar = Mid(LtrIn, i, 1)
        ReDim Preserve NumArray(i)
        NumArray(i) = Asc(TempChar) - 64
    Next
    HighPower = UBound(NumArray()) - 1
    
    For i = 1 To UBound(NumArray())
        TempNum = TempNum + (NumArray(i) * (26 ^ HighPower))
        HighPower = HighPower - 1
    Next
    colLetterToNum = TempNum
End Function
