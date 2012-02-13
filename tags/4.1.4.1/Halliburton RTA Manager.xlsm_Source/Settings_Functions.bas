Attribute VB_Name = "Settings_Functions"
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
      Call CMDline("/popUp", "", "Full-Screen mode: " & fsMode, 2)
  End Sub
  







