/*
 * * * Compile_AHK SETTINGS BEGIN * * *

[AHK2EXE]
Exe_File=C:\Dropbox\SVN\Halliburton RTA Manager\trunk\Include\RTA Sheet Tools.exe
Created_Date=1
Execution_Level=2
[VERSION]
Resource_Files=C:\Dropbox\SVN\Halliburton RTA Manager\trunk\Resource\tools.ico
Set_Version_Info=1
Company_Name=Halliburton - WellDynamics
File_Description=GUI application with convinient sheet tools for RTA Management Sheet
File_Version=1.0.0.20
Inc_File_Version=1
Internal_Name=RTA Sheet Tools
Original_Filename=RTA Sheet Tools
Product_Name=Source: AutoHotkey_L
Product_Version=1.1.2.0
Set_AHK_Version=1
[ICONS]
Icon_1=C:\Dropbox\SVN\Halliburton RTA Manager\trunk\Resource\tools.ico

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


;======================================================================================================
;GUI SETTINGS
;======================================================================================================
Gui, +AlwaysOnTop -caption +border
;~ Gui, -MaximizeBox
Gui, margin, 0, 0


;======================================================================================================
;MESSAGE HANDLER - CHANGE LV ROW COLOR WHEN SELECTED
;======================================================================================================
OnMessage("0x4E", "LVA_OnNotify")





;======================================================================================================
;Build GUI
;======================================================================================================
Gui, font, s10 w700
Gui, Color, White, White

;Window Header
Gui, add, Picture, x0 y-1 gTitleEvent section, %installDir%\Resource\RTA Sheet Tools Header 3d.png

;ListView Setup
Gui, Add, ListView, x195 ys+140 w310 r%RowCount% Checked AltSubmit vMyListView gMyListView Grid Section NoSort, |Column Header
Gui, Add, Button, xs+60 y+0  vOKButton,CLOSE

;AIS & LT's GroupBox
Gui, Add, GroupBox, x40 ys + 150 w113 h60 Section center, AIS && LT's
Gui, font, s10 w500
Gui, Add, Button, xs+5 ys+25 gHAISLT, Hide
Gui, Add, Button, xp+50 yp gSAISLT, Show
Gui, font, s10 w700

;PM Dates GroupBox
Gui, Add, GroupBox, xs yp+50 w113 h60 Section center, PM Dates
Gui, font, s10 w500
Gui, Add, Button, xs+5 ys+25 gHPMDates, Hide
Gui, Add, Button, xp+50 yp gSPMDates, Show	

;Make sure list view can receive color changes
LVA_ListViewAdd("MyListView")


;======================================================================================================
;RECEIVE VARIABLES FROM COMMAND LINE AND PUT INTO LIST VIEW
;======================================================================================================
;P1 variable parse for column names
loop, parse, p1, CSV, %A_Space%`t
{
		;Add value to the listview control
		LV_Add("", "", A_LoopField)
		
		;Set row color for non visible rows
		HideRowColor(A_Index)
}

;P2 variable parse for visible or hidden
loop, parse, p2, CSV, %A_Space%`t
{
	;Check box if visible
	LV_Modify(A_LoopField, "bold Check")
	
	;Set colors of rows for visible rows
	ShowRowColor(A_LoopField)
}

;Autosize the columns based on contents and headers
LV_ModifyCol(2, "AutoHdr")


;======================================================================================================
;GUI SHOW
;======================================================================================================
Gui, Show,, RTA Sheet Tools
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
if (x > 472 && x < 497 && y>0 && y<17)	;Minimize
	WinMinimize, RTA Sheet Tools
if (x>499 && x<540 && y>0  && Y< 17)		;Close
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
	
	If (A_EventInfo = 1 or A_EventInfo = 2 or A_EventInfo = 3 or A_EventInfo = 4)
	{
		LV_Modify(A_EventInfo, "-Select")
		LV_Modify(A_EventInfo, "Check")
		return
	}

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
			HideCol(XL, A_EventInfo, "h")
			
			LV_Modify(A_EventInfo, "-bold -Select")
		}
		
		;if row is unchecked then check / show it
		else if (InStr(ErrorLevel, "C", true))
		{
			LV_Modify(A_EventInfo, "Select")
			
			;Set colors of rows for visible rows
			ShowRowColor(A_EventInfo)

			;Show checked row
			HideCol(XL, A_EventInfo, "s")
			
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
			HideCol(XL, A_EventInfo, "h")
		}
		;If row is unchecked then check it / show it
		else 
		{
			LV_Modify(A_EventInfo, "Check")
			LV_Modify(A_EventInfo, "bold -Select")
			
			;Set colors of rows for visible rows
			ShowRowColor(A_EventInfo)

			;Show checked row
			HideCol(XL, A_EventInfo, "s")
		}
		return
	}
	
return

;======================================================================================================
;AIS & LT HIDE BUTTON EVENT
;======================================================================================================
HAISLT:
	LVA_ListViewAdd("MyListView")
	;~ SUOff()
	HideCol(XL,6, "h")
	HideRowColor(6)
	LV_Modify(6, "-Check")
	HideCol(XL,7, "h")
	HideRowColor(7)
	LV_Modify(7, "-Check")
	HideCol(XL,8, "h")
	HideRowColor(8)
	LV_Modify(8, "-Check")
	;~ SUOn()
return

;======================================================================================================
;AIS & LT SHOW BUTTON EVENT
;======================================================================================================
SAISLT:
	LVA_ListViewAdd("MyListView")
	;~ SUOff()
	HideCol(XL,6, "s")
	ShowRowColor(6)
	LV_Modify(6, "Check")
	HideCol(XL,7, "s")
	ShowRowColor(7)
	LV_Modify(7, "Check")
	HideCol(XL,8, "s")
	ShowRowColor(8)
	LV_Modify(8, "Check")
	;~ SUOn()
return

;======================================================================================================
;PM DATES HIDE BUTTON EVENT
;======================================================================================================
HPMDates:
	LVA_ListViewAdd("MyListView")
	SUOff()
	HideCol(XL, 13, "h")
	HideRowColor(13)
	LV_Modify(13, "-Check")
	HideCol(XL, 14, "h")
	HideRowColor(14)
	LV_Modify(14, "-Check")
	HideCol(XL, 15, "h")
	HideRowColor(15)
	LV_Modify(15, "-Check")
	HideCol(XL, 16, "h")
	HideRowColor(16)
	LV_Modify(16, "-Check")
	SUOn()
return

;======================================================================================================
;PM DATES SHOW BUTTON EVENT
;======================================================================================================
SPMDates:
	LVA_ListViewAdd("MyListView")
	SUOff()
	HideCol(XL, 13, "s")
	ShowRowColor(13)
	LV_Modify(13, "Check")
	HideCol(XL, 14, "s")
	ShowRowColor(14)
	LV_Modify(14, "Check")
	HideCol(XL, 15, "s")
	ShowRowColor(15)
	LV_Modify(15, "Check")
	HideCol(XL, 16, "s")
	ShowRowColor(16)
	LV_Modify(16, "Check")
	SUOn()
return

;======================================================================================================
;HIDE COLUMNS FUNCTION
;======================================================================================================
HideCol(XL, start, action = "h", end="")
{
	;~ end := end ? end : start
	if (action = "h")
		XL.sheets("RTA Manager").cells(1,start).EntireColumn.Hidden := true
	else if (action = "s")
		XL.sheets("RTA Manager").cells(1,start).EntireColumn.Hidden := false
}

;======================================================================================================
;SET COLORS FOR VISIBLE ROWS
;======================================================================================================
ShowRowColor(RowNum)
{
	LVA_SetCell("MyListView", RowNum, "0", "68E831", "Black")
}

;======================================================================================================
;SET COLORS FOR NOT VISIBLE ROWS
;======================================================================================================
HideRowColor(RowNum)
{
	LVA_SetCell("MyListView", RowNum, "0", "White", "Black")
}

;======================================================================================================
;GUI ESCAPE / CLOSE
;======================================================================================================
GuiEscape:
ButtonCLOSE:
GuiClose:
ExitApp

;======================================================================================================
;SCREEN UPDATING OFF
;======================================================================================================
SUOff()
{
	;~ XL.Application.ScreenUpdating = False
}

;======================================================================================================
;SCREEN UPDATING ON
;======================================================================================================
SUOn()
{
	;~ XL.Application.ScreenUpdating = true
}	
	
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