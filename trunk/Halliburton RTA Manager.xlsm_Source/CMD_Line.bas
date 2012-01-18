Attribute VB_Name = "CMD_Line"
' ___________________________________________________________________________________________________
' ***************************************************************************************************
' Module: ShellCalls
'
'   Functions that call external functions/programs via the shell. Most functions here
'   are related to the CMDline_Functions file packaged with the RTA Sheet (located in
'   the RTA Sheet's root dir\Include folder).
'
' About:
'   o Written by:    Rameen Bakhtiary
'   o Last Modified: 2012-01-16
' ___________________________________________________________________________________________________
' ***************************************************************************************************




' ___________________________________________________________________________________________________
' ===================================================================================================
' Sub: CMDline_Func
'
'   General function used to simplify shell calls to CMDline_Functions.exe.
'   Basically calls the program and passes each parameter that it receives as a parameter
'   to CMDline_Functions.exe
'
' Parameters:
'   cmdSwitch   -   The first parameter passed to CMDline_Functions. This is generally the
'                   cmd line switch.
'   param2      -   *(Optional)* Second parameter passed.
'   param3      -   *(Optional)* Third parameter passed.
'   param4      -   *(Optional)* Fourth parameter passed.
'   param5      -   *(Optional)* Fifth parameter passed.
'
' Last Modified: 2012-01-16
' ___________________________________________________________________________________________________
' ===================================================================================================
  Function CMDline_Func(cmdSwitch As String, Optional param2 = "", Optional param3 As String = "", Optional param4 = "", Optional param5 = "") As Double
      paramList = cmdSwitch & " " & """" & param2 & """" & " " & """" & param3 & """" & " " & """" & param4 & """" & " " & """" & param5 & """"
      CMDline_Func = Shell("""" & ThisWorkbook.Path & "\Include\CMDline_Functions.exe"" " & paramList, vbNormalFocus)
  End Function
  
  
  
  
  
  
  
  
