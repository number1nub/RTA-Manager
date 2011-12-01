VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RTA_Details 
   Caption         =   "Modify / View RTA Details"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13395
   OleObjectBlob   =   "RTA_Details.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RTA_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public saved As Boolean


'==================================================
'                  INITIALIZE
'==================================================
Private Sub UserForm_Initialize()
    Dim crw As Integer
    crw = ActiveCell.Row
    rtaNUm.Caption = "RTA " & ActiveCell.Value
    class = Cells(crw, getCol("Class"))
    desc = Cells(crw, getCol("Description"))
    comments = Cells(crw, getCol("Comments"))
    assignedTo = Cells(crw, getCol("Assigned To"))
    Department = Cells(crw, getCol("Current Status"))
    techrevdate = Cells(crw, getCol("Revised Due Date"))
    lab = Cells(crw, getCol("Lab Office"))
    rtatype = Cells(crw, getCol("Type"))
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




'==================================================
'          TRACK CHANGES - MARK AS UNSAVED
'==================================================
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



'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                       Double-Click RTA # - Open RTA in CWI
'
'       Description:    Open the RTA in CWI;  Uses external exe. Passes 6 digit RTA number as parameter

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub rtaNUm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    progP = ThisWorkbook.Path & "\Include\CMDline_Functions.exe"
    If Dir(progP) = "" Then Call MsgBox("Uh oh... An important file couldn't be found:  CMDline_Functions.exe" & vbCrLf & "" & vbCrLf & _
            "Without it, you cannot open RTAs directly in CWI. This file should be located" & Chr(10) & "in the Include folder within this worksheets directory. " & vbCrLf & vbCrLf & _
            "Running the installer file should solve this issue", vbCritical Or vbSystemModal, "--=[ WD RTA Sheet ]=--")
        Call Shell("""" & progP & """ " & Right(rtaNUm.Caption, 6), vbNormalFocus)
        Exit Sub
End Sub




'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                       Close Button -- Check if unsaved changes were made
'
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub closeME_Click()
    If saved = False And Application.Range("sheetviewmode") <> "PMT" Then
        ans = MsgBox("" & vbCrLf & "You made changes to this RTA that have not been saved." & vbCrLf & "" & vbCrLf & "Close without saving?", _
                vbYesNo Or vbExclamation Or vbSystemModal Or vbDefaultButton1, "     Changes Made!")
        If ans = vbNo Then desc.SetFocus: Exit Sub
    End If
    Unload Me
End Sub




'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 Reset Button
'
'

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub resetME_Click()
    Call UserForm_Initialize
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 uploadToCWI Button
'
'

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub uploadToCWI_Click()
    Dim rtasht As Worksheet
    Set rtasht = ThisWorkbook.Sheets("RTAimport")
    Application.ScreenUpdating = False
    rtasht.Visible = xlSheetVisible
    rtasht.Select
    
    ' Formatted RTA Number R00000XXXXXX
    '==================================
    tmp = "R00000" & Strings.Right(rtaNUm.Caption, 6)
    r = 1
    While Cells(r, 1) <> ""
        If Cells(r, 2) = tmp Then GoTo overwrite
        r = r + 1
    Wend
overwrite:
    Range("a" & r) = "Rta"
    Range("b" & r) = tmp
    'Replace all carriage returns with newlines and change 3 empy lines to 1
    '===========================================================
    Range("c" & r) = Replace(Replace(desc.Text, Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10), Chr(10)), Chr(13), "")
    Range("d" & r) = Replace(Replace(comments.Text, Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10), Chr(10)), Chr(13), "")
    
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
    Range("f" & r) = assignedTo
    Range("g" & r) = Department
    Range("h" & r) = techrevdate

    Application.DisplayAlerts = False

    rtasht.Select
    rtasht.Copy
    un = UserNameWindows
    ChDir "C:\documents and settings\" & un & "\my documents\"
    ActiveWorkbook.SaveAs Filename:="C:\documents and settings\" & un & "\my documents\rtaLoad.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
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
End Sub

