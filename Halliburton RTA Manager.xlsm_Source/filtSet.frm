VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} filtSet 
   Caption         =   "Group Filter Settings"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   OleObjectBlob   =   "filtSet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "filtSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image3_Click()

End Sub

' ===================================================================================================
'   Initialize the Filter Settings Userform
'
'       Get the lab office and filter settings, and then apply them to the sheet.
'
' ===================================================================================================
Private Sub UserForm_Initialize()
    prefix = Application.Range("cfilt")

    
    
    '===================================================
    '   Get current filter settings from  settings
    '   sheet ranges and create arrays containing the
    '   settings
    '===================================================
        
        
    '_______________
    '       RTA TYPE
    '
    Set tmprg = Application.Range(prefix & "type")
    i = 1
    rtaType.Clear
    For Each v In tmprg
        If v <> "" Then
             rtaType.AddItem tmprg(i)
            i = i + 1
        End If
    Next
    
    
    '_______________
    '       RTA CODE
    '
    Set tmprg = Application.Range(prefix & "code")
    i = 1
    rtacode.Clear
    For Each v In tmprg
        If v <> "" And v <> 0 Then
            rtacode.AddItem tmprg(i)
            i = i + 1
        End If
    Next
        
    
'    '____________
'    '       GROUP
'    '
'    Set tmprg = Application.Range(prefix & "group")
'    i = 1
'    Set tmprg2 = Application.Range(prefix & "groupDelim")
'    rtacode.Clear
'    For Each v In tmprg
'        If v <> "" And v <> 0 Then
'            group1.AddItem tmprg(i)
'            i = i + 1
'        End If
'    Next
'
    
    '________________
    '       RTA STATE
    '
    Set tmprg = Application.Range(prefix & "state")
    i = 1
    rtastate.Clear
    For Each v In tmprg
        If v <> "" Then
            rtastate.AddItem tmprg(i)
            i = i + 1
        End If
    Next

    
    '_____________________
    '       GET LAB OFFICE
    '
    Select Case prefix
    Case "pm"
        gTitle = "Permanent Monitoring Filter Settings"
    Case "fc"
    gTitle = "Flow Control Filter Settings"
    Case "di"
        gTitle = "Digital Infrastructure Filter Settings"
    Case "s"
        gTitle = "Software Filter Settings"
    End Select
End Sub

                
              



' ===================================================================================================
' _____ SUBMIT GUI _____
'
' Write the values from the form into the correct row/range on the settings sheet.
' After writing the settings to the sheet, call the main table sort routine.
' ===================================================================================================
Private Sub settingSubmit_Click()
    prefix = Application.Range("cFilt")
    
    '_______________
    '       RTA CODE
    '
    Dim fary1(1 To 10) As Integer
        For i = 0 To rtacode.ListCount - 1
            fary1(i + 1) = rtacode.List(i)
        Next i
    Application.Range(prefix & "code") = fary1
    
    '_______________
    '       RTA TYPE
    '
    Dim fary2(1 To 10) As String
        For i = 0 To rtaType.ListCount - 1
            fary2(i + 1) = rtaType.List(i)
        Next i
    Application.Range(prefix & "type") = fary2
    
    
    '________________
    '       RTA STATE
    '
    Dim fary3(1 To 10) As String
        For i = 0 To rtastate.ListCount - 1
            fary3(i + 1) = rtastate.List(i)
        Next i
    Application.Range(prefix & "State") = fary3
    
    
    '_____________
    '       GROUPS
    '
    Dim fAry4(1 To 10) As String
        For i = 0 To group.ListCount - 1
            fAry4(i + 1) = group.List(i)
        Next i
    Application.Range(prefix & "group") = fary3
     
    
    '________________________________________________________
    '       APPLY THE FILTER SETTINGS JUST INPUT TO THE TABLE
    '
    Unload Me
    sortBy (prefix)
End Sub












' ===================================================================================================
'   HANDLE GUI CONTROL ACTION: Add & remove filter items from listbox as clicked/added
'       by user.
'
' ===================================================================================================
 Private Sub CommandButton5_Click()
 
        '==================================================
        '        DOUBLE CLICK TO ADD/Remove FILTER
        '==================================================
        
        
        '_______________
        '       RTA CODE
        '
        Private Sub rtacode1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton7_Click
        End Sub
        Private Sub rtacode_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton8_Click
        End Sub
        
        
        '_________________
        '       RTA STATES
        '
        Private Sub rtastate1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton3_Click
        End Sub
        Private Sub rtastate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton4_Click
        End Sub
        
        '________________
        '       RTA TYPES
        '
        Private Sub rtatype1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton1_Click
        End Sub
        Private Sub rtaType_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton2_Click
        End Sub
        

        Private Sub CommandButton7_Click()
'RTA Code ADD
            If rtacode1 = "" Or IsNull(rtacode1) Then Exit Sub
            For i = 0 To rtacode.ListCount - 1
                If rtacode.List(i) = rtacode1 Then
                    rtacode1 = ""
                    Exit Sub
                End If
            Next
            rtacode.AddItem rtacode1
            rtacode1 = ""
        End Sub
        Private Sub CommandButton8_Click()
'RTA CodeREMOVE
            If rtacode = "" Or IsNull(rtacode) Then Exit Sub
            
            IND = rtacode.ListIndex
            rtacode.RemoveItem (IND)
        End Sub
        
        
        
        Private Sub CommandButton1_Click()
'RTA Type Add
            If rtatype1 = "" Or IsNull(rtatype1) Then Exit Sub
            
            For i = 0 To rtaType.ListCount - 1
                If rtaType.List(i) = rtatype1 Then
                    rtatype1.Selected(i) = False
                    Exit Sub
                End If
            Next
            rtaType.AddItem rtatype1
            rtatype1 = ""
        End Sub
        Private Sub CommandButton2_Click()
'RTA Type REMOVE
            If rtaType = "" Or IsNull(rtaType) Then Exit Sub
            
            IND = rtaType.ListIndex
            rtaType.RemoveItem (IND)
        End Sub
        
        
        
        Private Sub CommandButton3_Click()
'RTA State Add
            If rtastate1 = "" Or IsNull(rtastate1) Then Exit Sub
            
            For i = 0 To rtastate.ListCount - 1
                If rtastate.List(i) = rtastate1 Then
                    rtastate1.Selected(i) = False
                    Exit Sub
                End If
            Next
            rtastate.AddItem rtastate1
            rtastate1 = ""
        End Sub
        Private Sub CommandButton4_Click()
'RTA State REMOVE
            If rtastate = "" Or IsNull(rtastate) Then Exit Sub
            
            IND = rtastate.ListIndex
            rtastate.RemoveItem (IND)
        End Sub











' ===================================================================================================
' CLOSE BUTTON
'
' ===================================================================================================
Private Sub settingClose_Click()
    Unload Me
End Sub


