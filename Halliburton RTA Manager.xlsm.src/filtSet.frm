VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} filtSet 
   Caption         =   "Group Filter Settings"
   ClientHeight    =   9780
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
'==================================================
'                                               INITIALIZE FORM
'==================================================
Private Sub UserForm_Initialize()
    prefix = Application.Range("cfilt")

    '================= GET CURRENT SETTINGS ==================
    'RTA TYPE
    Set tmprg = Application.Range(prefix & "type")
    i = 1
    rtatype.Clear
    For Each v In tmprg
        If v <> "" Then
             rtatype.AddItem tmprg(i)
            i = i + 1
        End If
    Next
    'RTA CODE
    Set tmprg = Application.Range(prefix & "code")
    i = 1
    rtacode.Clear
    For Each v In tmprg
        If v <> "" And v <> 0 Then
            rtacode.AddItem tmprg(i)
            i = i + 1
        End If
    Next
    'RTA STATE
    Set tmprg = Application.Range(prefix & "state")
    i = 1
    rtastate.Clear
    For Each v In tmprg
        If v <> "" Then
            rtastate.AddItem tmprg(i)
            i = i + 1
        End If
    Next
    'PROD LT
    Set tmprg = Application.Range(prefix & "plt")
    pltMin = tmprg(1)
    pltMax = tmprg(2)
    'LT
    Set tmprg = Application.Range(prefix & "lt")
    ltMin = tmprg(1)
    ltMax = tmprg(2)
    
    'GUI Heading
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

                
              


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                           ADD and REMOVE  filter values
'                                               using the arrow buttons.  Ensure that no repeat values are allowed.
'

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
 Private Sub CommandButton5_Click()
 
         '==================================================
        '                                   DOUBLE CLICK TO ADD/Remove FILTER
        '==================================================
        ' RTA Code
        '==========
        Private Sub rtacode1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton7_Click
        End Sub
        Private Sub rtacode_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton8_Click
        End Sub
        
        
        ' RTA States
        '===========
        Private Sub rtastate1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton3_Click
        End Sub
        Private Sub rtastate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton4_Click
        End Sub
        
        ' RTA Types
        '===========
        Private Sub rtatype1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            Call CommandButton1_Click
        End Sub
        Private Sub rtatype_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
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
            
            For i = 0 To rtatype.ListCount - 1
                If rtatype.List(i) = rtatype1 Then
                    rtatype1.Selected(i) = False
                    Exit Sub
                End If
            Next
            rtatype.AddItem rtatype1
            rtatype1 = ""
        End Sub
        Private Sub CommandButton2_Click()
'RTA Type REMOVE
            If rtatype = "" Or IsNull(rtatype) Then Exit Sub
            
            IND = rtatype.ListIndex
            rtatype.RemoveItem (IND)
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


Private Sub settingClose_Click()
    Unload Me
End Sub



'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'                                                                                 Submit GUI
'
'                   and write changes to settings sheet
'

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub settingSubmit_Click()
    prefix = Application.Range("cFilt")
    
    'RTA CODE
    Dim fary1(1 To 10) As Integer
        For i = 0 To rtacode.ListCount - 1
            fary1(i + 1) = rtacode.List(i)
        Next i
    Application.Range(prefix & "code") = fary1
    'RTA TYPE
    Dim fary2(1 To 10) As String
        For i = 0 To rtatype.ListCount - 1
            fary2(i + 1) = rtatype.List(i)
        Next i
    Application.Range(prefix & "type") = fary2
    'RTA STATE
    Dim fary3(1 To 10) As String
        For i = 0 To rtastate.ListCount - 1
            fary3(i + 1) = rtastate.List(i)
        Next i
    Application.Range(prefix & "State") = fary3
    
    'PLT and LT; Must be passed as numbers for use
    'with data validation
    Dim p1 As Integer
    Dim p2 As Integer
    Dim l1 As Integer
    Dim l2 As Integer
    p1 = pltMin
    p2 = pltMax
    l1 = ltMin
    l2 = ltMax
    Application.Range(prefix & "plt")(1) = p1
    Application.Range(prefix & "plt")(2) = p2
    Application.Range(prefix & "lt")(1) = l1
    Application.Range(prefix & "lt")(2) = l2
    
    Unload Me
    sortBy (prefix)
End Sub


