#SingleInstance, Force
#NoEnv
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

;=======================================================================
;                  S C R I P T    S E T T I N G S        
;_______________________________________________________________________
;=======================================================================
	InstallerPath := "\\corp.halliburton.com\team\WD\public\Rameen Bakhtiary\setup files"
	UpdateInfoPath := InstallerPath "\update.ini"
	


;=================================================
;	Compare user's file version Beta build version
;=================================================
	
	;==============================
	;	Beta version
	;==============================
		IniRead, buildVersion, %UpdateInfoPath%, Beta, Version
		
		
	;==============================
	;	User version
	;==============================
		try RegRead, userVersion, HKCU, Software\Halliburton RTA Manager, Version
		catch 
			userVersion := "NONE"
	
	
	;==============================
	;	Compare versions
	;==============================
		if ((userVersion = "NONE") || (buildVersion >= userVersion)){
			try RunInstall(InstallerPath)
			catch e {
				;==============================
				;	Failed to install
				;==============================
				MsgBox, 4144, RTA Management Sheet BETA, % "Failed to complete installation.`n`n" e
				ExitApp
			}
		} else {
			MsgBox, Your Sheet is up to date!
			ExitApp
		}
	
	;==================================
	;	Log that user downloaded update
	;==================================
		formattime, DLtime,, hh:mm t, yyyy-MM-DD
		fileappend, `n - %dltime%`t`t`t%A_Username%`t`t`t%betaVersion%, %InstallerPath%\Internal\Beta Updater.txt
		sleep 100
ExitApp
		
		
;______________________________________________________________
;==============================================================
;	Function: RunInstall
;		Copies installer to user's temp dir & runs it.
;		Will throw a custom exception if an error occurrs.
;
;	Parameters:
;		_Path	-	Full path to MSI installer file
;==============================================================
RunInstall(_Path){
	;================================
	;	Copy the MSI to user Temp dir
	;================================
	try FileCopy, %_Path%\HRTA_%buildNum%.msi, %A_Temp%\HRTA_%buildNum%.msi
	catch
		throw "Unable to copy the installer from public server."

		
	;===================================
	;	Run the MSI and log the download
	;===================================
		try run, %A_Temp%\HRTA_%buildNum%.msi
		catch 
			throw "Error occurred while attempting to run installer file."
}





