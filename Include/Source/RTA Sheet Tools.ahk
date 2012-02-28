/*
 * * * Compile_AHK SETTINGS BEGIN * * *

[AHK2EXE]
Exe_File=C:\_.R.E.P.O.S._\Halliburton RTA Manager\Include\RTA Sheet Tools.exe
Created_Date=1
Execution_Level=2
[VERSION]
Resource_Files=C:\_.R.E.P.O.S._\Halliburton RTA Manager\Resource\tools.ico
Set_Version_Info=1
Company_Name=Halliburton - WellDynamics
File_Description=GUI application with convinient sheet tools for RTA Management Sheet
File_Version=1.0.0.70
Inc_File_Version=1
Internal_Name=RTA Sheet Tools
Original_Filename=RTA Sheet Tools
Product_Name=Source: AutoHotkey_L
Product_Version=1.1.5.6
Set_AHK_Version=1
[ICONS]
Icon_1=C:\_.R.E.P.O.S._\Halliburton RTA Manager\Resource\tools.ico

* * * Compile_AHK SETTINGS END * * *
*/

;_____________________________________________________________________
;---------------------------------------------------------------------
; TITLE:            R T A   S H E E T   T O O L S
;---------------------------------------------------------------------
; AUTHORS:
;	GUI & LAYOUT:		Sam T (acTennisKrazy)
;	FULL INTEGRATION:	Rameen B  (WSNHapps)
; DATE:			 		10/27/2011
; LANGUAGE:		 		English
; PLATFORM:		 		Created on Windows 7
; AHK-VERSION:	  		AutoHotkey_L
;---------------------------------------------------------------------
; DESCRIPTION:
; 
; External GUI that gives the RTA sheet additional functionality.
; Performs tasks  such as hiding rows and generating reports through
; and Excel COM interface.
;---------------------------------------------------------------------
 #NoEnv
 #SingleInstance, Ignore
 DetectHiddenWindows, On
 SetTitleMatchMode, RegEx 
 CoordMode, mouse, relative
 Functions()
;_____________________________________________________________________
;---------------------------------------------------------------------



AIS_COLUMNS = 


;======================================================================================================
; DETERMINE RTA SHEET INSTALL LOCATION
;======================================================================================================
InstallDir := Replace(RegRead("HKCU", "Software\Halliburton RTA Manager", "InstallDir"), """", "", "All")
if !(InstallDir)     ;Error with reg entry. Set working dir to parent dir of this script
	InstallDir = %a_scriptdir%\..






;======================================================================================================
;RegEx to match any possible window title of the excel window
;======================================================================================================
_Title = (.*(.*?(Excel).*?)?(RTA Manage(r|ment))(.*?(Excel).*?)?.*)
	
	
	
	
;======================================================================================================
;Tray menu icon from compiled resource
;======================================================================================================
if A_IsCompiled
	menu, tray, icon, % a_scriptfullpath, -159






;======================================================================================================
;INITIALIZE EXCEL OBJECT
;======================================================================================================
if ! WinExist(_Title) {
	sleep 500  
	MsgBox, The RTA Sheet is not open.
	IfWinNotExist,% _Title
		ExitApp
}
XL := Excel_Get()




;======================================================================================================
;CREATE OBJECT 'COL' THAT CONVERTS NUMBERS TO LETTERS. eg- col[3] CONTAINS "C"
;======================================================================================================
col := Object()
Loop, 26 {
	k:=chr(A_Index+64)
	col.insert(k)	
	i++
}







;======================================================================================================
;PARAMETER HANDLER - EXIT IF NOT RUN WITH PARAMETERS
;======================================================================================================
if 0=0
	ExitApp







;======================================================================================================
;ASSIGN PARAMETERS TO VARIABLES P1, P2,... PN
;======================================================================================================
loop, %0%
	p%a_index% := % %a_index%

	;Row count for listview control
	RowCount := 0
	loop, parse, p1, CSV, %A_Space%
		RowCount ++
	RowCount-=4




;======================================================================================================
;MESSAGE HANDLER - CHANGE LV ROW COLOR WHEN SELECTED
;======================================================================================================
OnMessage("0x4E", "LVA_OnNotify")






;======================================================================================================
;Build GUI
;======================================================================================================
buildGUI:
Gui, +AlwaysOnTop -caption +border

Gui, margin, 0, 0
Gui, font, s10 w500R
;~ Gui, font, s10 w700 c737373
Gui, Color, White, White
gui, Margin, 0, 10

;--- Title Image ---
Gui, add, Picture, y0 gTitleEvent section, %InstallDir%\Resource\RTA Sheet Tools Header.png


;--- Listview ---
gui, Font, cwhite
Gui, Add, ListView, x15 y+15 w230 BackgroundE0E0E0 r%RowCount% Checked AltSubmit vMyListView gMyListView Grid -Hdr NoSort Section, |Column Header


;--- AIS & LT's GroupBox ---
Gui, font, s10 w700 c737373
Gui, Add, GroupBox, x+30 ys-5 w115 h60 center Section, AIS && LT's
Gui, font, s10 w500

BUTTON_OPTS = W50 H30
Gui, Add, Button, xp+6 ys+20 %BUTTON_OPTS% gtoggleAIS, Show
Gui, Add, Button, x+2 yp %BUTTON_OPTS% gtoggleAIS, Hide


;--- PRESET Views GroupBox ---
gui, Font, w700
Gui, Add, GroupBox, xs-4 y+20 w125 h230 center Section, Preset Views
Gui, font, s10 w500

BUTTON_OPTS = w110 H30

Gui, Add, Button,  xs+7 ys+23 %BUTTON_OPTS% gpresetViews vpmtHide, PMT Mode
Gui, Add, Button,  xp y+5 %BUTTON_OPTS% gpresetViews vdeptHide, Edit Mode
Gui, Add, Button, xp y+5 %BUTTON_OPTS% gpresetViews vtrackRtaHide, RTA Tracking
Gui, Add, Button, xp y+5 %BUTTON_OPTS% gpresetViews vtrackTSGhide, TSG Mode
Gui, Add, Button, xp y+5 %BUTTON_OPTS% gpresetViews vrtaDatesHide, RTA Dates
Gui, Add, Button, xp y+5 %BUTTON_OPTS% gpresetViews vshowAll, Show All


gui, Add, Text, ys+140,

gosub, getColumnInfo


;______________________________________________
; 	GET POSITION OF EXCEL WINDOW AND CENTER GUI
;
GuiWidth := 325
WinGetPos, wx, wy, ww,,ahk_class XLMAIN    
guiX := Round(wx + ((ww/2) - (GuiWidth/2)))
Gui, Show, x%GuiX%, RTA Sheet Tools

;Make sure list view can receive color changes
LVA_ListViewAdd("MyListView")

return







;======================================================================================================
;RECEIVE VARIABLES FROM COMMAND LINE AND PUT INTO LIST VIEW
;======================================================================================================
;P1 variable parse for column names
getColumnInfo:
	loop, parse, p1, CSV, %A_Space%`t
	{
			if (A_LoopField="Priority" || A_LoopField="RTA" || A_LoopField="Customer" || A_LoopField="Description")
				continue
			
			StringReplace, val, A_LoopField, Production Lead time, Prod. LT
				
			;Add value to the listview control
			LV_Add("", "", val)
	}

		;Set row color for non visible rows
		HideRowColor(1, RowCount)

	;P2 variable parse for visible or hidden
	loop, parse, p2, CSV, %A_Space%`t
	{
		if (A_LoopField=1 || A_LoopField=2 || A_LoopField=3 || A_LoopField=4)
			continue
		
		thisRow := A_LoopField - 4
		
		;Check box if visible
		LV_Modify(thisRow, "bold Check")
		
		;Set colors of rows for visible rows
		ShowRowColor(thisRow)
	}

	;Autosize the columns based on contents and headers
	LV_ModifyCol(2, "AutoHdr")
return




;======================================================================================================
;GUI AND CONTROL RESIZING
;======================================================================================================
GuiSize:

	Anchor("MyListView", "w h")
	Anchor("TopText", "w", true)
	Anchor("OKButton", "y")
	
return


;===================================================
;						TITLE IMAGE EVENT
;===================================================
titleEvent:
MouseGetPos, x, y
if (x > 392 && x < 420 && y>8 && y<36)	;CLOSE BUTTON
	ExitApp
Else
	PostMessage, 0xA1, 2,,, A		;Move window when dragging header image
return





;======================================================================================================
;LIST VIEW  GUI EVENTS
;======================================================================================================
MyListView:

	;___________________________________
	;IF ONE OF THE FIRST 4 ROWS
	
	;~ If (A_EventInfo = 1 or A_EventInfo = 2 or A_EventInfo = 3 or A_EventInfo = 4)
	;~ {
		;~ LV_Modify(A_EventInfo, "-Select")
		;~ LV_Modify(A_EventInfo, "Check")
		;~ return
	;~ }




	;___________________________________
	;LISTVIEW CHECKBOX CLICKED EVENT
	
	if (A_GuiControlEvent = "I")
	{
		;If row is checked then uncheck / hide it
		if (InStr(ErrorLevel, "c", true))
		{
			LV_Modify(A_EventInfo, "Select")
			
			;Set row color for non visible rows
			HideRowColor(A_EventInfo)

			;Hide unchecked row
			HideCol(XL, A_EventInfo+4, "h")
			
			LV_Modify(A_EventInfo, "-bold -Select")
		}
		
		;if row is unchecked then check / show it
		else if (InStr(ErrorLevel, "C", true))
		{
			LV_Modify(A_EventInfo, "Select")
			
			;Set colors of rows for visible rows
			ShowRowColor(A_EventInfo)

			;Show checked row
			HideCol(XL, A_EventInfo+4, "s")
			
			LV_Modify(A_EventInfo, "bold -Select")
		}
	}
	
	


	;_____________________________________
	;LISTVIEW ROW DOUBLE CLICKED EVENT
	
	if (A_GuiControlEvent = "DoubleClick")
	{
		;If row is checked then uncheck / hide it
		if (LV_GetNext(A_EventInfo - 1, "C") = A_EventInfo)
		{
			LV_Modify(A_EventInfo, "-Check")
			LV_Modify(A_EventInfo, "-bold -Select")
			
			;Set row color for non visible rows
			HideRowColor(A_EventInfo)
			
			;Hide unchecked row
			HideCol(XL, A_EventInfo+4, "h")
		}
		;If row is unchecked then check it / show it
		else 
		{
			LV_Modify(A_EventInfo, "Check")
			LV_Modify(A_EventInfo, "bold -Select")
			
			;Set colors of rows for visible rows
			ShowRowColor(A_EventInfo)

			;Show checked row
			HideCol(XL, A_EventInfo+4, "s")
		}
		return
	}
	
return





;======================================================================================================
; TOGGLE AIS & LT COLUMNS
;======================================================================================================
toggleAIS:
	LVA_ListViewAdd("MyListView")
	
	xl.Range("aisHide").EntireColumn.Hidden := A_GuiControl = "Show" ? false : true
	
		
	%A_GuiControl%RowColor(2,4)
	
	LV_Modify(2, (A_GuiControl = "show" ? "" : "-") "check")
	LV_Modify(3, (A_GuiControl = "show" ? "" : "-") "check")
	LV_Modify(4, (A_GuiControl = "show" ? "" : "-") "check")
		
return






;======================================================================================================
;PM DATES HIDE BUTTON EVENT
;======================================================================================================
PresetViews:
	xl.cells.ENTIRECOLUMN.Hidden:=false
	xl.Range(A_GuiControl).ENTIRECOLUMN.Hidden:=true	
	Sleep 50
	ExitApp
return





;======================================================================================================
;HIDE COLUMNS FUNCTION
;======================================================================================================
HideCol(XL, start, action = "h", end="")
{
	global col
	End := end ? End : start
	xl.columns(col[start] ":" col[end]).ENTIRECOLUMN.Hidden := (action="h" ? true : false)
}



;======================================================================================================
;SET COLORS FOR VISIBLE ROWS
;======================================================================================================
ShowRowColor(RowNum, EndRow="")
{

	LoopCT := EndRow ? (EndRow-RowNum)+1 : 1
	RowNum--
	
	Loop, %LoopCT%
	{			
		whichRow := RowNum+A_Index
		LVA_SetCell("MyListView", whichRow, "0", "5B9CD7", "white")
	}
	;~ LVA_SetCell("MyListView", RowNum, "0", "5B9CD7", "White")
}



;======================================================================================================
;SET COLORS FOR NOT VISIBLE ROWS
;======================================================================================================
HideRowColor(RowNum, EndRow="")
{
	
	LoopCT := EndRow ? (EndRow-RowNum)+1 : 1

	Loop, %LoopCT%
	{		
		LVA_SetCell("MyListView", RowNum, "0", "E0E0E0", "black")
		RowNum++
	}
	;~ LVA_SetCell("MyListView", RowNum, "0", "E0E0E0", "black")
}



;======================================================================================================
;GUI ESCAPE / CLOSE
;======================================================================================================
GuiEscape:
GuiClose:
ExitApp





;======================================================================================================
;            INITIALIZE  COM  FUNCTIONS
;
; DESCRIPTION:    OBTAIN A COM CONNECTION WITH AN OPEN EXCEL INSTANCE 
; RETURN:      RETURNS AN ID VALUE THAT CAN BE USED AS "ACTIVEWORKBOOK"
;          IN EXCEL VBA
; AUTHORS & CREDITS:  TIDBIT (ON AHK FORUM)
;          JETHROW & SEAN 
;          HTTP://WWW.AUTOHOTKEY.COM/FORUM/VIEWTOPIC.PHP?T=67931
;======================================================================================================
	
Excel_Get(_WinTitle="ahk_class XLMAIN"){
  ControlGet, hwnd, hwnd, , Excel71, %_WinTitle%
  return, Excel_Acc_ObjectFromWindow(hwnd, -16).Application
}
Excel_Acc_Init(){
  Static  h
  If Not  h
	h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
}
Excel_Acc_ObjectFromWindow(hWnd, idObject = -4){
  Excel_Acc_Init()
  If  DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr", -VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
  Return  ComObjEnwrap(9,pacc,1)
}



;______________________ I N C L U D E   F I L E S _______________________________________
;-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#Include %a_scriptdir%\anchor.ahk
#Include %a_scriptdir%\lva.ahk