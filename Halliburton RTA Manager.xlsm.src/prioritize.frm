VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} prioritize 
   Caption         =   " ----==[ The Prioritizer ]==----"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   OleObjectBlob   =   "prioritize.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "prioritize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'                                                                                      SET  RTA  PRIORITY  GUI
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'==================================================
'      INITIALIZE GUI - GET EXISTING PRIORITY
'==================================================
Private Sub UserForm_Initialize()
    rtaNum.Caption = "RTA " & Cells(ActiveCell.Row, getCol("RTA"))
    priorty = ActiveCell.Text
End Sub


'==================================================
'             VALIDATE USER INPUT - Only allow numeric entry
'==================================================
Private Sub priorty_Change()
    If Not IsNumeric(priorty) And priorty <> "" Then
        MsgBox priorty & " is not a valid priority entry. Try again."
        priorty.SetFocus
        Exit Sub
    End If
End Sub

'==================================================
'                       CANCEL GUI
'==================================================
Private Sub CommandButton2_Click()
    setPriority = "Cancel"
    Unload Me
End Sub


'==================================================
'                                           SUBMIT GUI
'==================================================
Private Sub prioritySubmit_Click()
    setPriority = priorty
    Unload Me
End Sub

'==================================================
'                            ALLOW <ENTER> TO SUBMIT GUI
'==================================================
Private Sub priorty_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        Call prioritySubmit_Click
    End If
End Sub



