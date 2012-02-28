#SingleInstance, Force
#NoEnv
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

;=======================================================================
;                  S C R I P T    S E T T I N G S        
;_______________________________________________________________________
;=======================================================================
	BuildNum := "4.2.2"
	InstallerPath := "\\corp.halliburton.com\team\WD\public\Rameen Bakhtiary\setup files"
	


;================================
;	Copy the MSI to user Temp dir
;================================
	try copyFile(InstallerPath)
	catch e {
		msgbox % e
		ExitApp
	}
	
	
;===================================
;	Run the MSI and log the download
;===================================
	try run, %A_Temp%\HRTA_%buildNum%.msi
	catch {
		msgbox An error occurred while trying to run the MSI installer...
		ExitApp
	}
	sleep 200	
	ExitApp



;***********************************************************
; 	Function: copyFile
;		Copies MSI installer from public location to temp
;		dir. Throws a custom exception on error.
;***********************************************************
copyFile(_Path){
	try FileCopy, %_Path%\hrtam_%buildNum%.msi, %A_Temp%\HRTAM_%buildNum%.msi
	catch
		throw "An error occurred while trying to copy the installer from its public server.`n`n"
}