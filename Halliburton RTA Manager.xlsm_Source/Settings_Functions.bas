Attribute VB_Name = "Settings_Functions"
' ___________________________________________________________________________________________________
' ***************************************************************************************************
' Module: Settings_Functions
'
'   -
'
' About:
'   o Written by:    Rameen Bakhtiary
'   o Last Modified: 2012-01-16
' ___________________________________________________________________________________________________
' ***************************************************************************************************



'___________________________________________________________________________________________________
'***************************************************************************************************
'   Title: OTHER_FUNCTIONS
'---------------------------------------------------------------------------------------------------
'   Group: Overview
'       General overview of script features, functions & implementation
'
'       -
'
'   Group: About
'       Script file & source code information
'
'       - *Written By:*     Rameen Bakhtiary
'       - *Last Modified:*  1/10/2012
'---------------------------------------------------------------------------------------------------
'___________________________________________________________________________________________________
'***************************************************************************************************







'____________________________________________________________________________________________________
'====================================================================================================
' Function:     UserNameWindows
'
' Written by:   Rameen Bakhtiary
' Created on:   10/24/2011
' Description:
'               Returns the current user's Windows username
'
'====================================================================================================
Private Declare Function getUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Function UserNameWindows() As String
    Dim lngLen As Long
    Dim strBuffer As String
    Const dhcMaxUserName = 255
    strBuffer = Space(dhcMaxUserName)
    lngLen = dhcMaxUserName
    UserNameWindows = ""
    If CBool(getUserName(strBuffer, lngLen)) Then UserNameWindows = Left$(strBuffer, lngLen - 1)
End Function










'____________________________________________________________________________________________________
'====================================================================================================
' Function:     getCol
'
' Written by:   Rameen Bakhtiary
' Created on:   10/24/2011
' Description:
'               Returns the column number of a column given its header text
' Parameters:
'               header - String column header title of column whose number will be returned
'
'====================================================================================================
Function getCol(header As Variant) As Integer
    getCol = Range("Main[[#Headers],[" & header & "]]").Column
End Function








  
  
  
  

  
  
  
  

  
  
  
  
  
  
  '____________________________________________________________________________________________________
  '====================================================================================================
  ' Sub:          Toggle FullScreenMode_Click
  '
  ' Written by:   Rameen Bakhtiary
  ' Created on:   10/24/2011
  ' Description:
  '               Toggles the full screen mode display state on/off when checkbox is clicked.
  '
  '====================================================================================================
  Sub CheckBox6887_Click()
          
      'Show a pop-up notification of new full-screen mode
      '=====================================================
      If Application.Range("fullScreenMode") = True Then
          fsMode = "ON"
          Application.DisplayFullScreen = True
      Else
          fsMode = "OFF"
          Application.DisplayFullScreen = False
      End If
      Call CMDline_Func("/popUp", "", "Full-Screen mode: " & fsMode, 2)
  End Sub
  







