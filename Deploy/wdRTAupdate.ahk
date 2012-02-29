/*
 * * * Compile_AHK SETTINGS BEGIN * * *

[AHK2EXE]
Exe_File=\\corp.halliburton.com\team\WD\Public\RTA Manager\wdRTAupdate.exe
Created_Date=1
Execution_Level=2
[VERSION]
Resource_Files=C:\Dropbox\Halliburton RTA Manager\Include\Source\RTA Sheet Update.ahk_1.ico
Set_Version_Info=1
Company_Name=WSNHapps
File_Description=Checks for and, if required, deploys, updates for the Halliburton RTA Manager
File_Version=1.0.0.5
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
; TITLE: RTA Sheet Update
; ---------------------------------------------------------------------------------------------------
; 	Updater file for RTA Sheet. Should be placed in the 
; 	WD Public\RTA Manageer folder, and update.ini should also
; 	be in this location.
; ---------------------------------------------------------------------------------------------------
; 
; ABOUT:
; 	- Author: 	 Rameen B
;   - Modified:  2012-01-17
; ---------------------------------------------------------------------------------------------------
  #NoEnv
  #SingleInstance, Force
  SetWorkingDir %A_ScriptDir%  
  SetTitleMatchMode, RegEx
; ===================================================================================================




;=======================================================================
;                  S C R I P T    S E T T I N G S        
;_______________________________________________________________________
;=======================================================================
	;_________________________________________________
	; 		FULL PATH TO UPDATE SETTINGS/INFO INI FILE
	;
		update_INI := A_ScriptDir "\update.ini"


	;____________________________________________________________
	;		REGEX MATCH STRING FOR RTA MANAGER EXCEL WINDOW TITLE
	;
	;			Currently matches successfully when tested with
	;			all RTA sheet versions using Excel 2007 & 2010
	;			with any settings
	;
		ExcelTitle_RegEx := "(.*(.*?(Excel).*?)?(RTA Manage(r|ment))(.*?(Excel).*?)?.*)"
	 
	 
	;____________________________________________________________
	; 		REGEX STRING TO MATCH THE BETA VERSION INSTALLER FILE
	;		WITH THE FORMAT: HRTAM_<BUILD NUMBER>.MSI
	;
		BetaInstaller_Regex := "i)HRTAM_(?P<Version>\d\.\d\.\d(\.\d)?)\.msi"
	





;=================================================================
;    STORE INI INFO TO OBJECT
;=================================================================
	ini := new Ini(update_INI)






;=================================================================
;          G E T   U S E R ' S   S H E E T   V E R S I O N
;=================================================================

	;____________________________________________________
	; 		Get Sheet Version From Registry key Created
	;		by the RTA Sheet installer.
	;
	; 		If User's Version Isn't Found In The Registry
	;		Then Assume Userversion Is the "Previous" 
	;		version (latest version in CWI that didn't 
	;		write version info to the registry;
	;
	RegRead, userVersion, HKCU, Software\Halliburton RTA Manager, Version

	If (ErrorLevel || !userVersion)
		userVersion := ini.CWIrelease.previous
	;_____________________________________________________





;=================================================================
;	   C O M P A R E    V E R S I O N    W I T H   C W I
;=================================================================
	
	;_____________________________________________________
	; 		Compare user's version to the latest CWI 
	;		Release version. If a newer version is 
	;		released in CWI then prompt, close Excel
	;		and open the CWI Published Documents search
	;		page.
	;
	if (ini.CWIrelease.current > userVersion){
				
		msgbox, 4160, Halliburton RTA Manager, % "There is an updated version of the RTA Sheet released in CWI."
				  . "`n`nClick OK to open the published documents page and search for:`n9290-5103."		
		
		WinClose, % ExcelTitle_RegEx
		Sleep 100
		Process, Close, EXCEL.exe
		Sleep 100
		
		run, iexplore.exe http://cwiprod.corp.halliburton.com/cwi/ListPublishedFiles.jsp
		Sleep 100
		ExitApp
	}	
	;_____________________________________________________





;============================================================================================
;    I F   U S E R   O N   B E T A   L I S T ,   C H E C K   B E T A   V E R S I O N
;============================================================================================
	
	;______________________________________________
	; 		Get The List Of Beta Users From The INI
	;	 	And Check Against The Current User's 
	;		Windows Username; Exit if not in list.
	;
	BetaUserList := ini.BetaRelease.BetaUsers
	
	if A_UserName not contains %BetaUserList%
		ExitApp
	;______________________________________________
	
	
	
	;_____________________________________________________
	; 		Get The Folder In Which The Beta Installer Is
	;		Located 
	BetaInstaller_Dir := ini.BetaRelease.Folder
	;_____________________________________________________
	
	
	;______________________________________________________
	;		Get Build Version Number of Beta Release Using 
	;		Regex To Extract It From Its File Name.
	;		Compare To User's Version. If newer, then
	;		prompt, close Excel, copy the installer to 
	;		user's temp dir, and run it.
	;
	;		The version number of the beta build will be
	;		saved in variable "installerVersion" if the
	;		file is found.
	;
	Loop, A_ScriptDir betaInstaller_Dir "\*"
		
		If RegExMatch(A_LoopFileName, BetaInstaller_RegEx, installer)
		{
			
			;__________________________________________
			;*** Installer found; No updates needed ***
			if (userVersion >= installerVersion)
				ExitApp
			
			;________________________________________________________
			;*** Get release notes (if any) from releaseNotes.txt ***
			FileRead, releaseNotes, % A_ScriptDir BetaInstaller_Dir "\ReleaseNotes.txt"
			
			;________________________________________________________________
			;*** Prompt whether or not to install & display release notes ***
			msgbox, 36, Halliburton RTA Manager, % "A new build of the RTA Manager has been released."
					. "`nWould you like to install the latest version?`n`n"
					. "Notes on this release (if any):`n`n" ReleaseNotes
			
			;_____________________________
			;*** Declined update: Exit ***
			IfMsgBox, No
				ExitApp
			
			;_________________________
			;*** BEGIN BETA UPDATE ***
			TrayTip, Halliburton RTA Manager, Please wait while sheet updates....`n
			
			WinClose, % ExcelTitle_RegEx
			Sleep 100
			process, Close, EXCEL.exe
			Sleep 100
			
			FileDelete, % A_Temp "\RTAsheetBetaRelease.msi"
			Sleep 100
			FileCopy, % A_LoopFileFullPath, % A_Temp "\RTAsheetBetaRelease.msi", 1
			Sleep 150
			
			run, % A_Temp "\RTAsheetBetaRelease.msi"
			Sleep 200
			
			ExitApp		
		}
		




;===================================================================
;      I N C L U D E D   /   R E F E R E N C E   F I L E S
;===================================================================
#Include C:\Program Files\AutoHotkey\Lib\cIni.ahk






