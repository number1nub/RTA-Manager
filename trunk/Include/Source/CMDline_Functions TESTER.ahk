#SingleInstance, Force
#NoEnv
SetWorkingDir, %A_ScriptDir%  
SendMode, Input 
DetectHiddenText, on
DetectHiddenWindows, on
SetTitleMatchMode, 2
;~ Init := Functions()
;-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-


;~ InputBox,params, , Paramters:,,350,125
;~ If ErrorLevel or !params
	;~ ExitApp

Gui, Add, Edit, x72 y10 w90 h30 vp1, /popUp
Gui, Add, Edit, x72 y40 w300 h40 vtxt, 
Gui, Add, Edit, x72 y80 w300 h40 vtitle,
Gui, Add, Edit, x72 y120 w60 h30 vdur, 
Gui, Add, Edit, x72 y150 w300 h40 vrun, 
Gui, Add, Button, x142 y190 w100 h30 +Default, Submit
Gui, Add, Text, x22 y10 w50 h20 +Right, Param1:
Gui, Add, Text, x22 y50 w50 h20 +Right, Main Txt:
Gui, Add, Text, x32 y90 w40 h20 +Right, Title:
Gui, Add, Text, x22 y130 w50 h20 +Right, Duration:
Gui, Add, Text, x22 y160 w50 h20 +Right, Run:
; Generated using SmartGUI Creator for SciTE
Gui, Show, w393 h230, Untitled GUI
return

GuiClose:
ExitApp

enter::
buttonSubmit:
	gui, submit
	gui, destroy
	params := p1 (txt ? " "`"" txt "`""": " "`"""") (title ? " "`"" title "`""": " "`"""") (dur ? " "`"" dur "`""": " "`"""") (run ? " "`"" run "`""": " "`"""")
	run, cmdline_functions.ahk %params%
ExitApp







