/*
 * * * Compile_AHK SETTINGS BEGIN * * *

[AHK2EXE]
Exe_File=\\corp.halliburton.com\team\WD\public\RTA Manager\RTA Sheet Update.exe
Created_Date=1
Execution_Level=2
[VERSION]
Resource_Files=C:\Dropbox\Halliburton RTA Manager\Include\Source\RTA Sheet Update.ahk_1.ico
Set_Version_Info=1
Company_Name=WSNHapps
File_Description=Checks for and, if required, deploys, updates for the Halliburton RTA Manager
File_Version=1.0.0.2
Inc_File_Version=1
Internal_Name=RTA Sheet Update.exe
Legal_Copyright=WSNHapps
Original_Filename=RTA Sheet Update.exe
Product_Name=AutoHotkey_L
Product_Version=1.1.5.6
Set_AHK_Version=1
[ICONS]
Icon_1=%In_Dir%\RTA Sheet Update.ahk_1.ico

* * * Compile_AHK SETTINGS END * * *
*/

; ===================================================================================================
; TITLE: 	RTA Sheet Update
; ---------------------------------------------------------------------------------------------------
; Updater file for RTA Sheet. Should be placed in the 
; WD Public\RTA Manageer folder, and update.ini should also
; be in this location.
; ---------------------------------------------------------------------------------------------------
; 
; ABOUT:
; 	- Author: 	 Rameen B
;   - Modified:  2012-01-17
; ---------------------------------------------------------------------------------------------------
  #NoEnv
  #SingleInstance, Force
  SetWorkingDir %A_ScriptDir%  
  SetTitleMatchMode, 2
; ===================================================================================================




ini := new Ini("\\corp.Halliburton.com\team\WD\public\RTA Manager\update.ini")
MsgBox % ini.keys()





;=================================================================
;                  GET USER'S VERSION FROM  REGISTRY
;=================================================================
RegRead, userVersion, HKCU, Software\Halliburton RTA Manager, Version
If ErrorLevel || !userVersion
{
	
	;____________________________________________________________
	; 		User's Version Not In Registry =>> Compare "Previous"
	;		Cwi Version To Current And Check From New Release
	;
	if (ini.CWIrelease.current > ini.CWIrelease.previous){
        
        
        
    }
	
	
	IniRead, ForceUser, Setup Files\update.ini, ForceInstall Userlist, userlist, Err
	if ForceUser=Err
		ExitApp
	 

}

;=================================================================
;         GET CURRENT BUILD VERSION FROM INSTALLER'S FILENAME
;=================================================================
Loop, %A_ScriptDir%\Setup Files\*
	iF RegExMatch(A_LoopFileName, "i)HRTAM_(?P<Version>\d\.\d\.\d(\.\d)?)\.msi", installer)
		break



;=================================================================
;                  NO SETUP FILE FOUND IN THE FOLDER
;=================================================================
if !installerVersion
	ExitApp








