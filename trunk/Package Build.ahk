#NoEnv
#SingleInstance Force
SetBatchLines, -1
SetWorkingDir % A_ScriptDir


VersionFileRegEx := "i)[Version|Build][_|\s]?(?P<num>\d\.\d\.\d(\.\d)?)"


;_____________________
; Get the build number
;
loop, Build\Halliburton RTA Manager-SetupFiles\*
	if RegExMatch(A_LoopFileName, VersionFileRegEx, Build)
		goto BuildIT

;___________________________________________
; Prompt for build if version file not found
;
InputBox, buildNum,,`n  Enter the build number:
If ErrorLevel || StrLen(buildNum)<5
{
	msgbox Invalid build num. Goodbye.
	ExitApp
}


;===================================================
;			BEGIN  BUILDING
;===================================================
buildIT:
OutFile = RTA-Manager-Build-%buildNum%.zip
tempDir = buildRelease

;________________________________
; Delete the temp dir just incase
;
FileDelete, %outFile%
sleep 50
FileDelete, %tempDir%\*
sleep 50
FileRemoveDir, %tempDir%, 1
sleep 50

;_____________________________
; Extract the sheet VBA source
;
runwait, build\ExtractVBA.vbs `"%A_ScriptDir%\Halliburton RTA Manager.xlsm`"
sleep 200
	
	
;_________________________________
; Create the temp folder structure
;
FileCreateDir, %tempDir%\Source\VBA Source
sleep 50
FileCreateDir, %tempDir%\Source\AHK Source
sleep 50
FileCreateDir, %tempDir%\Bin
sleep 50


;___________________________
; Copy VBA source to folder
;
FileCopyDir, Halliburton RTA Manager VBA Source, %tempDir%\Source\VBA Source,1
sleep 50
FileDelete, Halliburton RTA Manager VBA Source\*
sleep 40
FileRemoveDir, Halliburton RTA Manager VBA Source, 1
sleep 40


;__________________________
; Copy AHK source to folder
;
FileCopy, Include\Source\*.ahk, %tempDir%\Source\AHK Source, 1
sleep, 50


;_______________
; Copy installer
;
FileCopy, Build\Halliburton RTA Manager-SetupFiles\*.msi, %tempDir%\Bin\Installer.msi, 1
sleep 50


;_______________
; Copy the sheet
;
FileCopy, Halliburton RTA Manager.xlsm, %tempDir%\Halliburton RTA Manager.xlsm, 1
sleep 50


;_______________
; Compress files
;
RunWait, 7z.exe a `"%outFile%`" %tempDir%\*, Hide
sleep 100


;______________________________
; Delete the temp setup folder
;
FileDelete, %tempDir%\*
sleep 50
FileRemoveDir, %tempDir%, 1
sleep 50


;~ msgbox Done













