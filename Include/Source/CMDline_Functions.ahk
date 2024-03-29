/*
 * * * Compile_AHK SETTINGS BEGIN * * *

[AHK2EXE]
Exe_File=C:\_.R.E.P.O.S._\Halliburton RTA Manager\Include\CMDline_Functions.exe
Created_Date=1
Execution_Level=2
[VERSION]
Resource_Files=C:\_.R.E.P.O.S._\Halliburton RTA Manager\Include\Source\tools.ico
Set_Version_Info=1
Company_Name=Halliburton - WellDynamics
File_Description=Functions, macros & scripts accessible via CMD line parameters. Supplements (and is required in order to run) the RTA Management Sheet.
File_Version=3.0.2.0
Inc_File_Version=0
Internal_Name=CMDline_Functions
Legal_Copyright=Rameen Bakhtiary - Halliburton|WellDynamics
Original_Filename=CMDline_Functions
Product_Name=Source - AutoHotkey_L
Product_Version=1.1.5.6
Set_AHK_Version=1
[ICONS]
Icon_1=%In_Dir%\tools.ico

* * * Compile_AHK SETTINGS END * * *
*/

;___________________________________________________________________________________________________
;***************************************************************************************************
;	Title: CMDline_Functions
;---------------------------------------------------------------------------------------------------
;	Group: Overview
;		General overview of script features, functions & implementation
;
;		Collection of functions & routines that run through the command line; helper script for
;       the RTA Manager.
;           
;       There are multiple commands/switches available; see the table below for each switch 
;       and it's corresponding function. This script performs no function if called without
;       arguments.
;
;    CMD Line Switches:
;
;       /checkRes                   -   Checks the screen resolution of the monitor that contains an                               
;                                       Excel window. If it's less than 1280 x 1024, prompts user to change
;                                       and will open the Windows Display Properties Settings dialogue.
;                                                 
;		/run <p2> [p3]              -   Runs p2;  if given, p3 will be passed as a parameter
;		            
;		/popUp <p2> [p3] [p4] [p5]  -   Display a flashing tooltip-like notification at bottom of screen.
;		                                p2 - Tool-tip main text
;	                                    p3 - Tool-tip title. (Default: WD RTA Management Sheet)                                                                                        
;	                                    p4 - Duration [seconds] to show popup. If p4 = "f" then popup
;	                                         will flash until it is clicked.
;	                                    p5 - If pop-up is clicked before timing out, this will be run
;	
;		/splash <p2>                -   Displays a splash image at the top of the screen with the specified text
;		                                given in p2
;            
;		/Load [p2]                  -	Opens a CWI page and gives instruction on how to load changes 
;		                                from an Excel sheet. If given - [p2] is copied to clipboard (to allow for
;		                                a path to be quickly pasted into the CWI dialogue).
;
;		/test [p2] [p3] [p4] [p5]   -	For debugging & testing purposes. Displays pop-up notification very
;		                                similar to /popup. See /popUp for parameter explainations (only
;		                                difference is that default title is "CMDline_Functions Debugging...")
;		                                                                
;		<p1> [p2]                   -	If a parameter is passed that does not begin with one of the above, 
;		                                it is assumed to be a CWI search; most commonly this is used to 
;		                                open an RTA in CWI (by double-clicking the Excel GUI title).             
;
;---------------------------------------------------------------------------------------------------
;	Group: About
;		Script file & source code information
;
;		- *Written By:* 	Rameen Bakhtiary
;		- *Last Modified:*	1/17/2012
;---------------------------------------------------------------------------------------------------
 #NoEnv
 #SingleInstance, Off
 DetectHiddenWindows, on
 DetectHiddenText, on
 CoordMode, mouse, relative
 SetWorkingDir, %A_ScriptDir%
 SetTitleMatchMode, 2
 SetBatchLines, -1
;___________________________________________________________________________________________________
;***************************************************************************************************




;===================================================
;         GET SHEET INSTALL PATH
;===================================================
try RegRead, installDir, HKCU, Software\Halliburton RTA Manager, InstallDir
catch
    installDir := A_MyDocuments "\Halliburton RTA Manager"


  
;===================================================
;       Tray Icon from compiled resources
;===================================================
If A_IsCompiled
    menu, tray, icon, % A_ScriptFullPath, -159


    

;==================================================
;            PARAMETER HANDLER
;==================================================
if 0>0
{   ;Assign parameters to variables p1, p2,... pN
    loop, %0%
        p%a_index% := % %a_index%
    if (StringLeft(p1, 1)<>"/")       ;Search CWI
        goto rtaSearch
    p1:=StringTrimLeft(p1, 1)
    if (IsFunc(p1))         ;Call function from param1
        %p1%()
    else if (IsLabel(p1))		;Call label from param1
        goto, %p1%
}
ExitApp





;=====================================================
;       /checkRes - Check screen resolution & Version
;=====================================================
checkRes:
    ;~ xl_Mon := WinGetMon(WinExist("ahk_class XLMAIN"))
    ;~ SysGet, xl_Coord, Monitor, %xl_Mon%
    ;~ SysGet, mainMon, Monitor, 1
    ;~ xRes := (xl_Mon > 1) ? (xl_CoordRight - mainMonRight) : xl_CoordRight
    ;~ yRes := xl_CoordBottom
    ;~ if ((xRes< 1280) || (yRes<1024)){       ;Resolution too low - prompt to change
        ;~ MsgBox, 4148, Halliburton Management Sheet, 
        ;~ (LTrim
        ;~ Your screen resolution is currently lower than what this sheet was designed to run on`n
        ;~ It is highly recommended that you change your resolution to at least 1280 x 1024 in order to avoid problems in performance.`n
        ;~ Open screen settings now? Press No to continue without changing (NOT RECOMMENDED)
        ;~ )
        ;~ IfMsgBox, yes
            ;~ RunWait, control desk.cpl`,`,3        
    ;~ }
ExitApp





;==================================================
;                 /LOAD - Load rtaLoad.xlsx to CWI
;==================================================
load:
    IfWinNotExist, Advanced Lookup -
    {
        run, iexplore.exe http://cwiprod.corp.halliburton.com/cwi/AdvLookup.jsp
        winwait, Advanced Lookup -        
        StatusBarwait,Done,,1,Advanced Lookup -
        sleep 200
        StatusBarwait,Done,,1,Advanced Lookup -
        sleep, 150
    }   
    sleep, 880
    winactivate, Advanced Lookup -        
    sleep 50
ExitApp





;'==================================================
;    POP UP - Display notify & show params
;'==================================================
popup:
	title := p1 = "test" ? "CMDline_Functions Debugging..." : p3 ? p3 : "`t`t"
	defaultTime := p4 ? p4 = "f" ? 0 : -p4 : -4.5
	bgOpts := "gc=white gt=off "
	titleOpts := "tc=A80000 ts=18 tw=750  tf=Century Gothic "
	txtOpts := "mc=black ms=13  mw=550  mf=Century Gothic "
    borderOpts := "bc=black bf=350  bt=off  bw=5 "
	picOpts := "iw=75 ih=75"
	behaveOpts := "st=800  si=1000  ac=clickPopup"
	notifyOpts := bgOpts titleOpts txtOpts borderOpts picOpts behaveOpts
    notify("`t" title "`t`t", p2 "`t`n", defaultTime, notifyOpts, A_ScriptDir "\..\Resource\Halliburton RTA Manager.png")
    if (p4 = "f"){
        sleep 10000
        ExitApp
    }
return    
    ;==================================
    ;  Action to take on pop-up click
    ;==================================
    clickPopup:
        if p5
            run, % p5
    ExitApp




;=================================================================
;            DISPLAY A 'PLEASE WAIT' SPLASH 
;=================================================================
Splash:
    ;___________________________________
    ; 	Get coordinates to center splash
    ;   in excel window
    ;
        splashWidth := 538
        WinGetPos, wx, wy, ww,,ahk_class XLMAIN    
        splashX := Round(wx + ((ww/2) - (splashWidth/2)))
	
    
    ;___________________________________
    ;   Set a backup time to destroy the
    ;   splash incase something happens
    ;   
        SetTimer, splashTimeout, 15000
        
    ;______________
    ; 	Show splash
    ;
        SplashImage, %A_ScriptDir%\..\Resource\Splash.png, CWwhite Zy0 zx0 x%SplashX% w%splashWidth% B1 FS16 WS600,,,RTA Management Splash

    ;______________________________________
    ; 	Wait for window popup from excel to
    ;   destroy splash, or until time out
    ;
        winwait, - Splash Off      
    
    splashTimeout:
        WinClose, RTA Manager - Splash Off    
        SplashImage, off   
ExitApp





;==================================================
;              /RUN --- Execute a run command
;==================================================
run:
    run, % p2 (p3 ? " " p3 : "")
ExitApp






;===================================================================================================
;   O P E N    R T A    I N    C W I  
;===================================================================================================
rtaSearch:

	; Check for passed value in INI file
    ;======================================
    IniRead, pID, \\corp.halliburton.com\team\wd\business development and technology\general\engineering public\PM App Logs\strMgmtLks\rtaid.ini, rtaids, %1%, Error
	
    
	; Open the MODIFY view of RTA found in INI
    ;============================================
	if pID != Error
	{	
        view := P2
        viewURL := "http://cwiprod.corp.halliburton.com/cwi/" 
					  . ((view = "n" || view = "nav" || view = "navigate") ? "Navigate.jsp?id=[pID]"
					  : (view = "sig" || view = "promote") ? "Navigate.jsp?dir=from&tableID=Approvals%23&id=[pID]"
					  : (view = "wu" || view = "where used") ? "Navigate.jsp?dir=to&id=[pID]"
					  : (view = "rta") ? "CreateModifyRta.jsp?id=[pID]"
                      : (view = "v") ? "View.jsp?id=[pID]"
					  : (view = "m" || view = "mod" || view = "modify") ? "Modify.jsp?id=[pID]"
					  : (view = "p" || view = "print") ? "View_noMenu.jsp?id=[pID]&flowPic=false&printFriendly=True"
                      : (view = "h") ? "History.jsp?id=[pID]"
					  : "StructureManagement.jsp?id=[pID]")
        StringReplace, url, viewURL, [pID], %pID%
        run, iexplore.exe %url%
        Sleep 50
        ExitApp
	}
    
    
	; Not found in INI -- search CWI
    ;===================================
    searchIT:
    IniRead, fieldXpos, %A_MyDocuments%\Halliburton RTA Manager\include\calibrationSettings.ini, fieldCoords, x, Error
    if (fieldxpos = "Error"){	;Calibration File not found
        ans := cmsgbox("CWI Search","Oh No!`n`nCWI Search calibration file not found. without it you`ncannot search CWI for objects that have not been indexed.`n`nWould you like to quickly calibrate now?","cblue", "Yes!|Not now", "resource\cwiicon.png")		
		if ans = Not now
			ExitApp
		
		; Calibrate CWI Search
		calibInstructions=
		(
		To calibrate CWI search:

    1.  Click "Calibrate" and a new CWI Advanced Lookup window will open.
         DO NOT click anything while waiting for it to open && fully load.
				
    2.  Once loaded, simply click in the "Search Text:" entry field 
         (Just as if you were going to enter a search)
		)
		ans := CMsgBox("CWI Search Calibration...",calibinstructions,"","Calibrate|Not now","I")
		if ans = Not now
			ExitApp
		;GO Calibrate
		gosub, calibrateCWI
        sleep 250
        goto searchIt
    }
    IniRead, sText_y, %A_MyDocuments%\Halliburton RTA Manager\include\calibrationSettings.ini, fieldCoords, ST
    IfWinNotExist, Advanced Lookup -
    {
        run, http://cwiprod.corp.halliburton.com/cwi/AdvLookup.jsp
        winwait, Advanced Lookup -        
        StatusBarwait,Done,,1,Advanced Lookup -
        sleep 200
        StatusBarwait,Done,,1,Advanced Lookup -
        sleep, 200
        StatusBarwait,Done,,1,Advanced Lookup -
        sleep, 100
    }        
    BlockInput, on 
    settimer, unBlock, 2000
    sleep, 200        
    WinActivate, Advanced Lookup -
    IfWinNotActive, Advanced Lookup -
        WinActivate, Advanced Lookup -
    sleep, 200
    MouseMove, %fieldXpos%, %sText_y%
    sleep 50
    SendInput, {click 3}
    Sleep, 100
    SendInput, {blind}%1%
    sleep, 50
    SendInput, {blind}{tab}{bs}{enter}
    BlockInput, off
ExitApp




;===================================================
;						CALIBRATE CWI SEARCH
;===================================================
calibrateCWI:
	; Open a CWI Adv. Lookup window
	run, iexplore.exe http://cwiprod.corp.halliburton.com/cwi/AdvLookup.jsp
    BlockInput, on
    winwait, Advanced Lookup -
    StatusBarwait,Done,,1,Advanced Lookup -
    sleep 200
    StatusBarwait,Done,,1,Advanced Lookup -
    qcalLabel:
    sleep, 250
    BlockInput, off
    WinActivate, Advanced Lookup -     
    IfWinNotActive, Advanced Lookup -
        WinActivate, Advanced Lookup -
    
    ;Wait for & record mouse-click
    keywait, lbutton, d
    MouseGetPos, quick_x, quick_y, win
    WinGetTitle, title, ahk_id %win%
	
    IfNotInString, title, Advanced Lookup -
    {	;Clicked outside of CWI Search window
		ans := CMsgBox("Quick CWI Search", "`nOops!  You didnt click in the CWI Advanced Search window...`n`nYou must click in the Search Tex field in the CWI window.", "", "Try it again|Forget it","E")
        if ans = Forget it
            ExitApp
        goto qcalLabel
    }
    sleep, 100
    
    ; Write to INI
    IniWrite, %quick_x%, %A_mydocuments%\Halliburton RTA Manager\include\calibrationSettings.ini, FieldCoords, x
    IniWrite, %quick_y%, %A_MyDocuments%\Halliburton RTA Manager\include\calibrationSettings.ini, FieldCoords, ST
    sleep, 300
	
	; Verify successful INI file creation
    if !(FileExist(A_MyDocuments "\Halliburton RTA Manager\include\calibrationsettings.ini")) {
        ans := CMsgBox("CWI Search Calibration", "`nUh oh... it looks like there was an error in writing to the calibration file.`n`nWould you like to give it another try?", "", "Try it again|No","E")
		if ans = No
			ExitApp
        goto qcalLabel
    }
    
	;DONE
    ans:=CMsgBox("CWI Search Calibration", "Success! `n`nSettings file was created.", "","Complete Setup","I")
    sleep, 50
return




;==================================================
;                           Safety unlock incase of error
;==================================================
unblock:
    settimer, unblock, off
    BlockInput, off
return        




;==================================================
;						WinGetMon Function
;
;		Description:		Return the index of the monitor containing
;	a given window (from HWND of window).
;	Default index is 1
;==================================================
WinGetMon(windowHandle){
   IfWinNotExist, ahk_id %windowHandle%	;Ensure "windowHandle" is open
      return
   monitorIndex := 1		;Default monitor index to 1
   VarSetCapacity(monitorInfo, 40)
   NumPut(40, monitorInfo)   
   if (monitorHandle := DllCall("MonitorFromWindow", "uint", windowHandle, "uint", 0x2)) 
		&& DllCall("GetMonitorInfo", "uint", monitorHandle, "uint", &monitorInfo) 
	{  monitorLeft   := NumGet(monitorInfo,  4, "Int")
		monitorTop    := NumGet(monitorInfo,  8, "Int")
		monitorRight  := NumGet(monitorInfo, 12, "Int")
		monitorBottom := NumGet(monitorInfo, 16, "Int")
		workLeft      := NumGet(monitorInfo, 20, "Int")
		workTop       := NumGet(monitorInfo, 24, "Int")
		workRight     := NumGet(monitorInfo, 28, "Int")
		workBottom    := NumGet(monitorInfo, 32, "Int")
		isPrimary     := NumGet(monitorInfo, 36, "Int") & 1
      SysGet, monitorCount, MonitorCount
      Loop, %monitorCount%
      {	SysGet, tempMon, Monitor, %A_Index%
         if ((monitorLeft = tempMonLeft) and (monitorTop = tempMonTop)
            and (monitorRight = tempMonRight) and (monitorBottom = tempMonBottom))  ; Compare location to determine the monitor index.
         {	monitorIndex := A_Index
            break
         }
      }
   }
   return %monitorIndex%
}




;=======================================================================
;		AHK Command Wrapper Functions
;		Description:			Functions for certain AHK commands that have "output vars"
;		Original License:		Version 1.41 <http://www.autohotkey.net/~polyethene/#functions>
;		Modified by:				Rameen Bakhtiary
;		Modification Date:	9/12/2011
;=======================================================================
IsEqual(var, val) {
	if (var=val)
		return 1
} 
IniRead(Filename, Section, Key, Default = "Error") {
	IniRead, v, %Filename%, %Section%, %Key%, %Default%
	Return, v
}
RegRead(RootKey, SubKey, ValueName = "") {
	RegRead, v, %RootKey%, %SubKey%, %ValueName%
	Return, v
}
StringLeft(ByRef InputVar, Count) {
	StringLeft, v, InputVar, %Count%
	Return, v
}
StringRight(ByRef InputVar, Count) {
	StringRight, v, InputVar, %Count%
	Return, v
}
StringTrimLeft(ByRef InputVar, Count) {
	StringTrimLeft, v, InputVar, %Count%
	Return, v
}
StringTrimRight(ByRef InputVar, Count) {
	StringTrimRight, v, InputVar, %Count%
	Return, v
}
Replace(ByRef InputVar, SearchText, ReplaceText = "", All = ""){
    StringReplace, v, InputVar, %searchText%, %ReplaceText%, %All%
    Return, v
}



;===================================================
;		Custom MsgBox Function
;
;		Author:		Danny Ben Shitrit (aka Icarus)
;
CMsgBox( title, text, textOpts="", buttons="", icon="", owner=0 ) {
  Global _CMsg_Result  
  GuiID := 9      ; If you change, also change the subroutines below
  StringSplit Button, buttons, |
  If( owner <> 0 ) {
    Gui %owner%:+Disabled
    Gui %GuiID%:+Owner%owner%
  }
  Gui %GuiID%:+Toolwindow +AlwaysOnTop -theme -caption +border
  gui %guiid%:color, white, red
  if icon not contains png,ico,bmp,jpg
  {MyIcon := ( icon = "I" ) or ( icon = "" ) ? 222 : icon = "Q" ? 24 : icon = "E" ? 110 : icon
    Gui %GuiID%:Add, Picture, Icon%MyIcon% , Shell32.dll
  }
  else
    Gui %GuiID%:Add,Picture, , %icon%
  gui, %GuiID%:font, s10 w600
  Gui %GuiID%:Add, Text, %textOpts% x+12 yp r8 section , %text%
  gui, %GuiID%:font, s8.5 w600
  Loop %Button0% 
    Gui %GuiID%:Add, Button, % ( A_Index=1 ? "x+12 ys " : "xp y+3 " ) . ( InStr( Button%A_Index%, "*" ) ? "Default " : " " ) . "w120 gCMsgButton", % RegExReplace( Button%A_Index%, "\*" )
  Gui %GuiID%:Show,,%title%
  Loop 
    If( _CMsg_Result )
      Break
  If( owner <> 0 )
    Gui %owner%:-Disabled
  Gui %GuiID%:Destroy
  Result := _CMsg_Result
  _CMsg_Result := ""
  Return Result
}
9GuiEscape:
9GuiClose:
  _CMsg_Result := "Close"
Return
CMsgButton:
  StringReplace _CMsg_Result, A_GuiControl, &,, All
Return


