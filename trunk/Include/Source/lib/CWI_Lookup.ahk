;---------------------------------------------------------------------------------------------------
;	Title: CWI Search Functions Wrapper
;	
;		Set of functions that perform actions relating to searching CWI and/or opening
;		object pages in CWI.	
;---------------------------------------------------------------------------------------------------



;***************************************************************************************************
;	Function: CWI_Lookup
;		Main entry-point function for performing searches and other actions in CWI. 
;
;		Call CWIsearch with your search text as the parameters and it takes care of everything
;		else.
;
;	Note:
;		The method used to perform searches implements the custom view notation implemented in
;		the PM App Launcher tool, CWI Search Bar, where a ">" separates the search string from
;		a view identifier code as shown in the example below.
;
;		(Start code AHK)
;			; Open 9290-9868 in CWI in Navigate view
;			CWIsearch("9290-9868>n")
;		(End)
;
;		See the <CWI Object Class> documentation for more details on CWI search options and features.
;
CWI_Lookup(sTxt, tabbed = ""){
	
	;_________________________________
	; Special trigger keyword searches
	;
		if !(sTxt){
			open_lookup()
			return
		} 
		else if (sTxt = "qcal"){
			Calibrate()
			return
		}
	;______________________________
	; Save search txt for history
	;
		add_History(sTxt)
	;______________________________
	; Auto-open from saved CWI info
	;		
		cView := get_CustomView(sTxt)
		oCwi := new CWI
		Url := oCwi.Url(sTxt, cView)	
		if (Url){
			Run, iexplore.exe %Url%
			ErrorLevel:=""
			return
		}
		
	;____________________________________
	; Manual send-strokes Advanced Lookup
	;
		Loop {
			cal := get_Calib()
			if ErrorLevel
			{	
				;Attempt to migrate if not already tried
				RegRead, triedMigration, HKCU, Software\PM App Launcher\CWI Search Bar, didMigrate
				if triedMigration != 1
					Migrate_Settings()
				Calibrate()
				If ErrorLevel
					ExitApp
			}
		} until !(ErrorLevel)
				
		win := open_lookup()	
		SetKeyDelay, 10
		SetMouseDelay, 10
		CoordMode, mouse, relative
		
		MouseMove, % cal.x, % cal.y
		Click 3
		Sleep, 50
		Send, %sTxt%{Tab}
		Sleep 50
		Send, {Backspace}
		send, {Enter}

		if !(win.existed || tabbed){
			sleep 2000
			winclose, % "ahk_id " win.id
		}
}





;***************************************************************************************************
;	Function: get_Calib
;		Returns an object containing the X and Y calibration coordinates of the "Search Text" field in 
;		a CWI Advanced Lookup window.
;
;	Return Value:
;		An array containing 2 elements indexed as "x" and "y" is returned containing the calibration
;		values (as integers). For example-
;			> If one or both of the calibration values are not found then ErrorLevel is set to 1; ErrorLevel 
;		is set to 0 otherwise.
;	
get_Calib(){
	RegRead, xC, HKCU, Software\PM App Launcher\CWI Search Bar\Calibration, xCoord	
	RegRead, yC, HKCU, Software\PM App Launcher\CWI Search Bar\Calibration, yCoord	
	ErrorLevel := (xC = "" || yC = "") ? 1 : ""	
	calib := {x:xC, y:yC}
	return calib
}





;***************************************************************************************************
;	Function: Calibrate
;		Handles the process of obtaining the mouse-click coordinates required by the search
;		function. The user is prompted with instructions, and upon completing the calibration
;		the values are stored in the registry and *ErrorLevel* is set to 0.	If for any reason 
;		the values don't get saved properly into the registry then *ErrorLevel* is set to 1.
;
;	Parameters:
;		noPrompt -	*(Optional)* If this value is passed as 1 then the instruction prompts
;					will be skipped. Not recommended to be used when calling the function
;					since it is always wise to give instruction to user; this is mostly 
;					utilized by the function itself in cases where a "retry" requires calling
;					itself and re-prompting would be annoying.
;
Calibrate(noPrompt = ""){
	CoordMode, mouse, Relative
	
	;______________________________________
	; PROMPT THE USER WITH SOME INSTRUCTION
	; (UNLESS noPrompt=1)
	;
		if !(noPrompt){
				
			InStruction_Txt=
			(LTrim, RTrim
				Quick and Easy Calibration!!
			
				`t1.) Wait for the CWI Advanced Lookup window to open/load
			
				`t2.) Click in the `"Search Text`" field -- That's it!
			)
			
			msgbox, 4161, CWI Search Calibration, % InStruction_Txt			
			IfMsgBox Cancel
				return
		}
	;________________________________________
	; OPEN A  CWI ADV. LOOKUP WINDOW AND WAIT
	; UNTIL IT IS LOADED & ACTIVE
	;
		open_Lookup()
		
	;____________________________________
	; WAIT FOR AND RECORD THE MOUSE CLICK
	;
		Retry:
		KeyWait, LButton, d
		MouseGetPos, mx, my, win
		Sleep, 75
		WinGetTitle, title, ahk_id %win%
		
	;______________________________________________________________
	; USER CLICKED IN AN INVALID LOCATION. PROMPT & TRY AGAIN/ABORT
	;
		If !instr(title, "Advanced Lookup")
		{
			msgbox, 4149, , Ooops!`n`nIt looks like you clicked in the wrong window--`n`nRemember`, just click in the `"Search Text`" field to calibrate!
		
			;______________________________________________________
			; CLICKED RETRY - LOOP BACK TO BEGINNING OF CALIBRATION
			;
			IfMsgBox Retry
			{
				WinClose, Advanced Lookup -
				Sleep 100
				goto ReTry
			}
			;_________________
			; OTHERWISE, ABORT
			;
			else
			{
				msgbox, 4096, , Aborting calibration...
				ErrorLevel := 1
				return
			}
		}				
	
	;____________________________________________
	; RECEIVED VALID INPUT -- WRITE VALUES TO REG
	;		
		WinClose, Advanced Lookup -
		writeCalib(mx, my)		
		ErrorLevel := ErrorLevel ? 1 : ""
	
		If !(ErrorLevel)
			msgbox, CWI Search calibration successfully saved!			
	return
}




;***************************************************************************************************
;	Function: writeCalib
;		Writes the given values to the proper location in the registry. The function performs
;		some validation by ensuring the Reg values exist and are equal to the passed values,
;		and sets ErrorLevel to 0 if all is good.
;
;	Parameters:
;		mx	- X-Coordinate to write to registry
;		my	- Y-Coordinate to write to registry
;
writeCalib(mx, my){	
	;_________________________
	; WRITE VALUES TO REGISTRY
	;
		RegWrite, REG_SZ, HKCU, Software\PM App Launcher\CWI Search Bar\Calibration, xCoord, %mx%		
		Sleep 50
		RegWrite, REG_SZ, HKCU, Software\PM App Launcher\CWI Search Bar\Calibration, yCoord, %my%
		Sleep 50
	;___________________
	; VALIDATE THE WRITE
	;
		RegRead, test, HKCU, Software\PM App Launcher\CWI Search Bar\Calibration, yCoord
		if (test = "" || test != my){
			msgbox, 4144, , There was an issue while trying to write calibration data`nto your registry. Try re-installing the PM App Launcher package to fix the issue.`n`nPlease report this issue if it persists.
			ErrorLevel:=1
			return
		}
	ErrorLevel:=""
	return
}




;***************************************************************************************************
;	Function: get_CustomView
;
;	Parameters:
;		sTxt	-	The search text entry to save into the history as "last searched"
;
add_History(sTxt){
	RegWrite, REG_SZ, HKCU, Software\PM App Launcher\CWI Search Bar\History, lastSearch, %sTxt%
}





;***************************************************************************************************
;	Function: get_CustomView
;		Returns the last searched item from the CWI Search Bar history
;
get_History(){	
	RegRead, retVal, HKCU, Software\PM App Launcher\CWI Search Bar\History, lastSearch
	return retVal
}





;***************************************************************************************************
;	Function: get_CustomView
;
;		Returns the custom view identifier if found in the given text. The function also modifies
;		the search string passed by removing the identifier so that only the actual search text 
;		remains.
;
;	Parameters:
;		sTxt	- 	A string in which to look for a CWI view modifier. Passed
;					ByRef so that the view modifier can be removed from the string when found.
;
get_CustomView(ByRef sTxt){
	if RegExMatch(sTxt, ">(?P<view>\w+)?$", cust){
		StringTrimRight, sTxt, sTxt, % (StrLen(custview) + 1)
		return custView
	}
}




;***************************************************************************************************
;	Function: open_Lookup
;
;		Opens/activates a CWI Advanced Lookup window. Waits until the window is fully loaded
;		and active before returning. Returns the HWND of the IExplore window launched.
;
;	Return Value:
;		Returns the window handle (HWND) of the Advanced Lookup IE window
;
open_Lookup(){	
	SetTitleMatchMode, 2
	didExist := 1
	
	;_______________________________________________________
	; 	CHECK FOR & ACTIVATE EXISTING ADVANCED LOOKUP WINDOW
	;
	If not winID := WinExist("Advanced Lookup -")
	{				
		didExist:=""
		;________________________________________________________
		;OPEN AN ADV. LOOKUP WINDOW AND WAIT FOR IT TO FULLY LOAD
		;
		Run, iexplore.exe http://cwiprod.corp.halliburton.com/cwi/AdvLookup.jsp
		WinWaitActive, Advanced Lookup -
		IfWinNotActive, Advanced Lookup -
			WinActivate, Advanced Lookup -
		winID := WinExist("A")	
		Sleep, 400	
		StatusBarWait,Done,,1,Advanced Lookup -
		Sleep 200
		StatusBarWait,Done,,1,Advanced Lookup -
		Sleep 200
		StatusBarWait,Done,,1,Advanced Lookup -
		Sleep 200
	}
	WinActivate, ahk_id %wid%
	sleep 350
	winInfo := {id:winID, existed:didExist}
	return winInfo
}







;***************************************************************************************************
;	Function: Migrate_Settings
;		This function goes through the user's settings and files in attempt to gather missing 
;		settings/data that changed location from previous versions. For example, coordinate 
;		values previously stored in an INI file are now being stored and looked for in the registry;
;		when this function is called it searches all previous setting locations and writes values
;		to the updated location if found.
;
;	Note:
;		Function marks in the user's registry noting that migration was attempted to prevent wasting
;		time/memory checking what has already been checked.
;
Migrate_Settings(){
	
	;Record that migration has been attempted
	RegWrite, REG_SZ, HKCU, Software\PM App Launcher\CWI Search Bar, didMigrate, 1
		
	;Try to set calibration from old calib. locations
	if checkSettings(A_MyDocuments "\PM App Launcher\Include\CWI SB\callibrationSettings.ini")
		goto settingsFound
	else if checkSettings(A_MyDocuments "\Halliburton RTA Manager\Include\callibrationSettings.ini")
		goto settingsFound
	else if checkSettings(A_MyDocuments "\Halliburton RTA Manager\Include\calibrationSettings.ini")
		goto settingsFound
	else
	{
		ErrorLevel := 1
		return
	}		
	settingsFound:
	ErrorLevel := ""
}



;***************************************************************************************************
;	Function: CheckSettings
;		Sets CWI calibration data in registry if found in the given file. If the file doesn't exist 
;		or if the file doesn't contain valid CWI callibration data, then the function returns
;		an empty value & ErrorLevel is set to 1; if calibration data is found & set, Errorlevel is
;		set to 0.
;
;	Parameters:
;		_filePath	-	The full path of the file in which to look for CWI callibration settings	
;
checkSettings(_filePath){
	If fileexist(_filePath)
	{
		IniRead, oldX, %_filePath%, fieldCoords, x, Err
		IniRead, oldY, %_filePath%, fieldCoords, ST, Err
		if ((oldX <> "Err" && oldX > 50) && (oldY <> "Err" && oldY > 50))
		{
			;Found calibration
			writeCalib(oldX, oldY)
			ErrorLevel := ""
			return 1
		}
	}
	ErrorLevel := 1
}







;***************************************************************************************************
;	 Class:	Cwi
;		CWI Object class used to quickly get, add and change CWI object information. Also allows 
;		methods for quickly obtaining and opening an objects CWI URL for instant viewing in CWI. 
;		On initialization of a new instance, the full contents of the CWI object data files are stored
;	 	to variables and are accessible.		
;
;	Parameters:
;		The class instantiator does accept an *optional* parameter representing the root path of the 
;		CWI Object info INI files. The current default is set to the Eng. Pub. "pm app logs\strMgmtLks"
;	 	folder. 
;
;	Remarks:
;		Upon creation of a new instance of the object, the following global variables are available:
;		- [instance name].allSAPnums - Contains list of all Ref number / SAP number relations recorded.
;		- [instance name].ALLids - Contains list of all object number / CWI jspID number relations recorded.
;
;		ERRORLEVEL:
;
;		ErrorLevel is set to 1 if there was a problem constructing the class.
;
;	Example: 
;		(Start Code AHK)
;			cwi := new cwi()
;			If ErrorLevel
;			{	MsgBox An error occurred while creating an instance of the class... `n`nAborting.
;				ExitApp
;			}
;			
;			MsgBox % cwi.url("9290-2658")
;			
;			MsgBox % cwi.url("744677")
;		(End)
;
class CWI
{		
	idINI_path := ""
	sapINI_path := ""
	rtaINI_path := ""


	;***************************************************************************************************
	;	Function: __New
	;		Constructor function for new class instances. 
	;
	;		Upon creation of a new instance of the class, 3 globally accessible variables are assigned--
	;		allSAPnums, allIDs and allRTAs. each of these is a string containing a list of all the value
	;		entries within the respective INI files.
	;
	__New(FolderPath = "\\corp.halliburton.com\team\WD\Business Development and Technology\General\Engineering Public\PM App Logs\strMgmtLks")
	{				
		this.INI_FolderPath := RegExReplace(FolderPath, "i)\\$")		
		;__________________________________
		; 		Full paths to the ini files
		;
		this.IDini_path := this.INI_FolderPath "\objectids.ini"
		this.SAPini_path := this.INI_FolderPath "\partSAPnums.ini"
		this.RTAini_path := this.INI_FolderPath "\rtaID.ini"
				
		;_____________________________________________________
		; 	Read full contents of each INI into objects in mem
		;
		this.IDs := new Ini(this.idINI_path)
		this.sapNums := new Ini(this.sapINI_path)
		this.rta := new Ini(this.rtaINI_path)						
		

		this.allIDs := FileRead(this.idINI_path)
		this.allSAPnums := FileRead(this.sapINI_path)
		this.allrtas := FileRead(this.rtaINI_path)
		
		ErrorLevel := this.IDs.Sections() ? 0 : 1
	}
	
	

	



		
	;***************************************************************************************************
	;
	; Function: get_SAPnum
	;
	;		Returns the SAP number of an object, given its reference number.
	;
	; Parameters:
	;		refNum - Reference number of the object who SAP number is to be returned.
	; Return Value: 
	;		Returns SAP number of object if found; otherwise, returns 0.
	; Remarks:
	;		If incorrect parameter format is received (i.e.- a reference number is not passed) function returns "Error: Format"
	; Related:
	;		get_REFnum, get_ID
	;
	get_SAPnum(refNum){
		if ! RegExMatch(refNum, "(\d{4}-\d{4})"){
			ErrorLevel:=1
			return
		}
		ErrorLevel=
		return this.sapNums.sapnums[refNum]		
	}
	




	;***************************************************************************************************
	;
	; Function: get_REFnum
	;
	;		Returns the REF number of an object given any type of input.
	;
	; Parameters:
	;		fromNum - Either an SAP or Reference number of an object; if an SAP number is passed, a REF number will be found and returned.
	;			If a reference number is passed, it will simply be returned. 
	; Return Value: 
	;		Returns reference number of object. If an SAP number is passed and not found, returns 0.
	;
	; Related:
	;		get_SAPnum, get_ID
	;
	get_REFnum(fromNum){
		if RegExMatch(fromNum, "\d{4}-\d{4}"){
			ErrorLevel=
			return fromNum
		}
		if this.find_Part(fromNum){
			ErrorLevel=
			return substr(this.allSAPnums, this.find_Part(fromNum) - 10, 9)
		}
	}




	;***************************************************************************************************
	;
	; Function: get_ID
	;
	;		Returns the CWI jsp ID number of an object given any type of input.
	;
	; Parameters:
	;		fromNum - Either an SAP or Reference number of an object whose ID will be returned.
	;
	; Return Value: 
	;		If found, returns the ID of the given object, otherwise returns empty. Also, ErrorLevel is set to either
	;		"rta," "task" or "other" if the ID is found. ErrorLevel is set to 0 if the part isn't found.
	;		
	; Remarks:
	;		Function will convert from SAP to reference number and vise-versa in order to check all possible resources for
	;			the ID number. If not found, return is 0.
	; Related:
	;		get_SAPnum, get_REFnum, url
	;
	get_ID(fromNum){
		if instr(this.allids, fromnum){	 ;In objectIDs
			ErrorLevel = other
			return this.ids.partids[fromNum]
		}
		else if (this.ids.partids[this.get_REFnum(fromNum)]){	 ;In objectIDS (after converting SAP)
			ErrorLevel = Other
			return this.ids.partids[this.get_REFnum(fromNum)]
		}
		else if (this.get_rtaID(fromNum)){	
			ErrorLevel := ErrorLevel
			return this.get_rtaID(fromNum)
		}
		ErrorLevel:=0
	}



	;***************************************************************************************************
	;
	; Function: get_rtaID
	;
	;		Returns the CWI jsp ID number of an RTA or Task object
	;
	; Parameters:
	;		rtaNum - The number of the RTA/Task to look up (with the leading letter & zeros removed)
	;
	; Return Value: 
	;		Returns the jsp ID of the given object if found. If the found object was a task, then 
	;		ErrorLevel is set to "task" -- if the object was an RTA then ErrorLevel="rta" -- otherwise
	;		ErrorLevel is blank (if number not found).
	;
	get_rtaID(rtaNum){
		if instr(this.allrtas, rtanum){
			if (this.rta.rtaids[rtaNum]){
				ErrorLevel=rta
				return this.rta.rtaids[rtaNum]
			}else{
				ErrorLevel=task
				return this.rta.taskids[rtaNum]
			}
		}
		ErrorLevel=1
		return
	}




	;***************************************************************************************************
	;
	; Function: get_taskID
	;
	;		Exactly identical to <get_rtaID>. This function is only added for consistency and 
	;		familiarity.
	;
	; Parameters:
	;		rtaNum - The number of the RTA/Task to look up (with the leading letter & zeros removed)
	;
	; Return Value: 
	;		Returns the jsp ID of the given object if found. If the found object was a task, then 
	;		ErrorLevel is set to "task" -- if the object was an RTA then ErrorLevel="rta" -- otherwise
	;		ErrorLevel is blank (if number not found).
	;
	get_taskID(rtaNum){
		if instr(this.allrtas, rtanum){
			if (this.rta.rtaids[rtaNum]){
				ErrorLevel=rta
				return this.rta.rtaids[rtaNum]
			}else{
				ErrorLevel=task
				return this.rta.taskids[rtaNum]
			}
		}
		ErrorLevel=1
		return
	}



	;***************************************************************************************************
	; Function: URL
	;		Returns the complete URL to an objects page in CWI in a specified view.
	;
	; Parameters:
	;		fromNum - Either an SAP or Reference number of an object whose CWI page URL is to be returned. 
	;		view - (optional) Specifies the CWI page view to open the object in. See Remarks for more info. Default view is 
	;			Structure Management.
	; Return Value: 
	;		Returns a  full URL string containing the specified objects jsp ID and the address to the specified CWI view.
	;		Returns empty if not found.
	;
	;		Note that ErrorLevel is set to 0 if url not found; if the object's ID is found, then ErrorLevel is
	;		set to either "other," "rta" or "task"
	;
	; Remarks:
	;		View modes:
	;			n		-	Opens in Navigate view
	;			sig		-	Opens the View Signitures page
	;			wu		-	Opens Where Used for part
	;			rta		-	Opens the Create/Modify view for RTAs
	;			m		-	Opens Modify view
	;			p		-	Opens printer friendly page
	;			h		-	Opens part's history
	;
	URL(fromNum, view=""){
				
		;____________________
		; 	GET THE PART'S ID
		;
		ObjectID := this.get_id(fromNum)
	
		;_______________________
		; 		ID NOT FOUND....
		;
		if !(ObjectID){
			ErrorLevel=
			return
		}
		
		;__________________________________________
		; 	DETERMINE OBJECT TYPE FROM GET_ID'S ERRORLEVEL
		;	AND SET THE VIEW MODE
		;
		Type := ErrorLevel
		view := view ? view
			  : ErrorLevel = "rta" ? "rta"
		      : ErrorLevel = "task" ? "sig"
			  : ""			  			  
			  
		;______________
		; 		GET URL
		;
		viewURL := "http://cwiprod.corp.halliburton.com/cwi/" 
			   . ((view = "n" || view = "nav" || view = "navigate") ? "Navigate.jsp?id=[pID]"
			   : (view = "sig" || view = "promote") ? "Navigate.jsp?dir=from&tableID=Approvals%23&id=[pID]"
			   : (view = "wu" || view = "where used") ? "Navigate.jsp?dir=to&id=[pID]"
			   : (view = "h" || view = "history") ? "History.jsp?id=[pID]"
			   : (view = "rta") ? "CreateModifyRta.jsp?id=[pID]"
			   : (view = "m" || view = "mod" || view = "modify") ? "Modify.jsp?id=[pID]"
			   : (view = "p" || view = "print") ? "View_noMenu.jsp?id=[pID]&flowPic=false&printFriendly=True"
			   : "StructureManagement.jsp?id=[pID]")
		
		StringReplace, viewURL, viewURL, [pID], %ObjectID%
		ErrorLevel := Type
		return viewURL
	}
						





	;***************************************************************************************************
	;	Function: find_Part
	;	
	;		Returns the position of a given number in the SAP number INI file.
	;
	;	Parameters:
	;		fromNum - SAP (or REF) num to find in the SAP ini file. Usually an SAP number, as this function
	;				is mainly used in getting a reference number from an SAP number.
	;
	;	Returns:
	;		The starting position of the SAP number, if found. Otherwise returns blank & ErrorLevel is set to 1.
	;
	find_Part(fromNum){
		if instr(fromNum, "-"){
			ErrorLevel:=1
			return
		}
		ErrorLevel=
		return instr(this.allSAPnums, fromNum)
	}	






	;***************************************************************************************************
	;
	; Function: inFile
	;
	;		Used to determine weather or not a given object number is found anywhere in a specified file.
	;
	; Parameters:
	;		fromNum - Either an SAP or Reference number of an object to look for in index file
	;		file - (optional) Specifies the file to look in. Default is "SAP," or the file containing relations between reference
	;			numbers and SAP numbers. Pass anything into the file parameter and the part ID file will be searched instead.
	; Return Value: 
	;		Returns true if the part is found in the file, or empty otherwise. Errorlevel is set
	;		to 1 if not found and 0 otherwise.
	;
	; Example:
	;		if inFile("9290-6676", "id")
	;			msgbox Found it!
	;
	inFile(fromNum, file="sap"){
		found := instr(file = "sap" ? this.allSAPnums : this.allIDs, fromNum)
		ErrorLevel := found ? 0 : 1
		return found
	}	







	;***************************************************************************************************
	;
	; Function: add_SAPnum
	;
	;		Add a new entry of REF num = SAP num to the data index files.
	;
	; Parameters:
	;		refNum - Reference number of object to add
	;		SAPnum - SAP number of part to add
	;
	;
	;
	; Related:
	;		add_ID
	;
	add_SAPnum(refNum, SAPnum){
		if !(refNum || sapnum){
			ErrorLevel:=1
			return
		}
		ErrorLevel=
		this.sapNums.sapnums[refNum] := SAPnum    
		this.sapNums.save(this.SAPini_path)
	}


	;***************************************************************************************************
	;
	; Function: add_ID
	;
	;		Add a new entry of Object num = jsp ID to the data index files.
	;
	; Parameters:
	;		refNum - Reference number of object to add
	;		pID - CWI jsp ID of the part to be saved
	;
	;
	;
	; Related:
	;		add_SAP
	;
	add_ID(fromNum, pID){
		if !(fromNum || pID){
			ErrorLevel:=1
			return
		}
		ErrorLevel=
		this.ids.partids[fromNum] := pID
		this.ids.save(this.idINI_path)
	}
	
	
	
	;***************************************************************************************************
	;
	; Function: add_rtaID
	;
	;		Add a new entry of RTA num = jsp ID to the data index files.
	;
	; Parameters:
	;		rtaNum - RTA/Task number of object to add
	;		pID - CWI jsp ID of the part to be saved
	;
	;
	add_rtaID(rtaNum, pID){
		if !(rtaNum || pID){
			ErrorLevel:=1
			return
		}
		ErrorLevel=			
		this.rta.rtaids[rtaNum] := pID
		this.rta.save(this.rtaINI_path)
	}
	
	
	
	
	
	
	
	;***************************************************************************************************
	;
	; Function: add_taskID
	;
	;		Add a new entry of Task num = jsp ID to the data index files.
	;
	; Parameters:
	;		rtaNum - RTA/Task number of object to add
	;		pID - CWI jsp ID of the part to be saved
	;
	;
	add_taskID(taskNum, pID){
		if !(taskNum || pID){
			ErrorLevel:=1
			return
		}
		ErrorLevel=
		this.rta.taskids[taskNum] := pID
		this.rta.save(this.rtaINI_path)
	}
	
	
	
	

		
	;***************************************************************************************************
	;
	; Function: del_SAP
	;
	;		Removes a specified entry of REF num = SAP num from the data index file given any object number.
	;
	; Parameters:
	;		fromNum - Either the reference number or SAP number of the part  whose data is to be removed from the file.
	;
	;
	; Remarks:
	;		Function will return an empty value on successfully finding and deleting an entry; 1 is returned otherwise.
	; Related:
	;		del_ID
	;
	del_SAP(fromNum){
		if !instr(this.allSapnums, fromNum){
			ErrorLevel:=1
			return
		}
		ErrorLevel=
		this.sapNums.delete("sapNums", fromNum)
		this.sapNums.save(this.sapINI_path)
	}



	;***************************************************************************************************
	;
	; Function: del_ID
	;
	;		Removes a specified entry of object num = jsp ID from the data index file given any object number.
	;
	; Parameters:
	;		fromNum - Either the reference number or SAP number of the part  whose data is to be removed from the file.
	;
	;
	; Remarks:
	;		Function will return an empty value on successfully finding and deleting an entry; 1 is returned otherwise.
	; Related:
	;		del_SAP
	;
	del_ID(fromNum){
		if !InStr(this.allIDs, fromNum){	;Not found in IDs -- break
			ErrorLevel:=1
			return
		}
		ErrorLevel=
		this.ids.delete("partIDs", fromNum)
		this.ids.save(this.idINI_path)
	}

}




;===================================================
;			File read helper function
;===================================================
FileRead(File){
	fileread, v, %file%
	return v
}





/* ***************************************************************************************************
	Class: Ini
		Class that allows efficient and simple methods to manipulate INI files.
		
	Original Author & Link:
		- By: (zzzooo10)
		- Forum Post: <http://www.autohotkey.com/forum/viewtopic.php?p=462061#462061>


  ***************************************************************************************************
 */
class Ini
{
		; Loads ini file.
	__New(File, Default = "") {
		If (FileExist(File)) and (RegExMatch(File, "\.ini$"))
			FileRead, Info, % File
		Else
			Info := File
		Loop, Parse, Info, `n, `r
		{
			If (!A_LoopField)
				Continue
			If (SubStr(A_LoopField, 1, 1) = ";")
			{
				Comment .= A_LoopField . "`n"
				Continue
			}
			RegExMatch(A_LoopField, "(?:^\[(.+?)\]$|(.+?)=(.*))", Info) ; Info1 = Seciton, Info2 = Key, Info3 = Value\
			If (Info1)
				Saved_Section := Trim(Info1), this[Saved_Section] := { }, this[Saved_Section].__Comments := Comment, Comment := ""
			Info3 := (Info3) ? Info3 : Default
			If (Info2) and (Saved_Section)
				this[Saved_Section].Insert(Trim(Info2), Info3) ; Set the section name withs its keys and values.
		}
	}
	
	__Get(Section) {
		If (Section != "__Section")
			this[Section] := new this.__Section()
	}
	
	class __Section
	{ 
		__Set(Key, Value) {
			If (Key = "__Comment")
			{
				Loop, Parse, Value, `n
				{
					If (SubStr(A_LoopField, 1, 1) != ";")
					{
						NewValue .= "; " . A_LoopField . "`n"
						Continue
					}
					NewValue .= A_LoopField . "`n"
				}
				this.__Comments := NewValue
				Return NewValue
			}
		}
		
		__Get(Name) {
			If (Name = "__Comment")
				Return this.__Comments
		}
	
	}
	
	; Renames an entire section or just an individual key.
	Rename(Section, NewName, KeyName = "") { ; If KeyName is omited, rename the seciton, else rename key.
		Sections := this.Sections(",")
		If Section not in %Sections%
			Return 1
		else if ((this.HasKey(NewName)) and (!KeyName)) ; If the new section already exists.
			Return 1
		else if ((this[Section].HasKey(NewName)) and (KeyName)) ; If the section already contains the new key name.
			Return 1
		else if (!this[Section].HasKey(KeyName) and (KeyName)) ; If the section doesn't have the key to rename.
			Return 1
		else If (!KeyName)
		{
			this[NewName] := { }
			for key, value in this[Section]
				this[NewName].Insert(Key, Value)
			this[NewName].__Comment := this[Section].__Comment
			this.Remove(Section)
		}
		Else
		{
			KeyValue := this[Section][KeyName]
			this[Section].Insert(NewName, KeyValue)
			this[Section].Remove(KeyName)
		}
		Return 0
	}
	
	; Delete a whole section or just a specific key within a section.
	Delete(Section, Key = "") { ; Omit "Key" to delete the whole section.
		If (Key)
			this[Section].Remove(Key)
		Else
			this.Remove(Section)
	}
	
	; Returns a list of sections in the ini.
	Sections(Delimiter = "`n") {
		for Section, in this
			List .= (this.Keys(Section)) ? Section . Delimiter : ""
		Return SubStr(List, 1, -1)
	}
	
	; Get all of the keys in the entire ini or just one section.
	Keys(Section = "") { ; Leave blank to retrieve all keys or specify a seciton to retrieve all of its keys.
		Sections := Section ? Section : this.Sections()
		Loop, Parse, Sections, `n
			for key, in this[A_LoopField]
				keys .= (key = "__Comments" or key = "__Comment") ? "" : key . "`n"
		Return SubStr(keys, 1, -1)
	}
	 
	; Saves everything to a file.
	Save(File) { 
		Sections := this.Sections()
		loop, Parse, Sections, `n
		{
			NewIni .= (this[A_LoopField].__Comments)
			NewIni .= (A_LoopField) ? ("[" . A_LoopField . "]`n") : ""
			For key, value in this[A_LoopField]
				NewIni .= (key = "__Comments" or key = "__Comment") ? "" : key . "=" . value . "`n"
			NewIni .= "`n"
		}
		FileDelete, % File
		FileAppend, % SubStr(NewIni, 1, -1), % File
	}
	
}











