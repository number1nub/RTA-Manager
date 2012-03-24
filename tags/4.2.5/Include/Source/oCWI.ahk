/* ***************************************************************************************************
 *  	Title: CWI Object Class
 * ***************************************************************************************************
 */
 


/*  ===================================================================================================
 *  CLASS:	Cwi
 *  ~~~~~~~~~~~~~~~~~~~~~
 * 	CWI Object class used to quickly get, add and change CWI object information. Also allows 
 * 	methods for quickly obtaining and opening an objects CWI URL for instant viewing in CWI. 
 * 	On initialization of a new instance, the full contents of the CWI object data files are stored
 *  to variables and are accessible.		
 *
 *	PARAMETERS:
 *  ~~~~~~~~~~~~~~~~~~~~~
 * 	The class instantiator does accept an *optional* parameter representing the root path of the 
 * 	CWI Object info INI files. The current default is set to the Eng. Pub. "pm app logs\strMgmtLks"
 *  folder. 
 *
 * 	REMARKS:
 *  ~~~~~~~~~~~~~~~~~~~~~
 * 	Upon creation of a new instance of the object, the following global variables are available:
 * 	- [instance name].allSAPnums - Contains list of all Ref number / SAP number relations recorded.
 * 	- [instance name].ALLids - Contains list of all object number / CWI jspID number relations recorded.
 *
 * 	ERRORLEVEL:
 *  ~~~~~~~~~~~~~~~~~~~~~
 *	ErrorLevel is set to 1 if there was a problem constructing the class.
 *
 *	EXAMPLE: 
 *  ~~~~~~~~~~~~~~~~~~~
 * ;		(Start Code)
				cwi := new cwi()
				If ErrorLevel
				{	MsgBox An error occurred while creating an instance of the class... `n`nAborting.
					ExitApp
				}
				
				MsgBox % cwi.url("9290-2658")
				
				MsgBox % cwi.url("744677")
				
 *
 *  ===================================================================================================
 */
class CWI{
		
		
	idINI_path := ""
	sapINI_path := ""
	rtaINI_path := ""
	
	;=================================================================
	;              CONSTRUCTOR FOR NEW INSTANCE OF CWI CLASS
	;=================================================================
	__New(FolderPath = "\\corp.halliburton.com\team\WD\Business Development and Technology\General\Engineering Public\PM App Logs\strMgmtLks"){		
		
		
		this.INI_FolderPath := RegExReplace(FolderPath, "i)\\$")
		
		
		;***************************************************************************************************
		;   Group: Global Variables
		;       Global variables created when class is constructed
		;
		;   	IDini_path 	- Full path to the object ID ini file
		;  		SAPini_path - Full path to the CWI part SAP number ini file
		;   	allSAPnums 	- A string list of every <REF Num - SAP Num> entry in the ini file
		;   	allIDs 		- A string list of every <REF Num - ID Num> entry in the ini file
		;
		;   Group: Functions
		;___________________________________________________________________________________________________
		;***************************************************************************************************	
		
		;__________________________________
		; 		FULL PATHS TO THE INI FILES
		;
		this.IDini_path := this.INI_FolderPath "\objectids.ini"
		this.SAPini_path := this.INI_FolderPath "\partSAPnums.ini"
		this.RTAini_path := this.INI_FolderPath "\rtaID.ini"
				

		;_____________________________________________________________________
		; 		CREATE AN INI OBJECT FOR EACH OF THE FILES USING THE INI CLASS
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
	;
	; Function: URL
	;
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
	;			"n" "nav" or "navigate"		-	Opens in Navigate view
	;			"sig" or "promote"				- 	Opens in View Signitures view
	;			"wu" or "where used"			-	Opens in Where Used view
	;			"m" "mod" or "modify"			-	Opens in RTA Modify view
	;			Blank or any other input		-	Opens in Structure Management (default)
	; Related:
	;		get_ID
	;
	URL(fromNum, view=""){
		
		
		;===================================================
		;			GET THE PART'S ID
		;===================================================
		ObjectID := this.get_id(fromNum)
		
		
		;_______________________
		; 		ID NOT FOUND....
		;
		if !(ObjectID){
			ErrorLevel=
			return
		}
		
		Type := ErrorLevel
		
		
		;____________________________________________
		; 		SET THE DEFAULT VIEW IF NOT SPECIFIED
		;
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
						
					;~ ;-- Try RTAs & Tasks ---
		
		;~ rID := this.get_rtaID(fromNum)
		;~ If !(ErrorLevel){  ;-- Not found
			;~ ErrorLevel=1
			;~ return				
		;~ }
		
		;~ ;----  Set the default views for tasks/Rtas ----
		;~ ;-----------------------------------------------
		;~ view := view ? view
			 ;~ : (ErrorLevel = "task") ? "sig"
			 ;~ : "rta"
		;~ ;-----------------------------------------------
		
		;~ viewURL := !(rID) ? "" 
					  ;~ : "http://cwiprod.corp.halliburton.com/cwi/" 
					  ;~ . ((view = "n" || view = "nav" || view = "navigate") ? "Navigate.jsp?id=[pID]"
					  ;~ : (view = "sig" || view = "promote") ? "Navigate.jsp?dir=from&tableID=Approvals%23&id=[pID]"
					  ;~ : (view = "wu" || view = "where used") ? "Navigate.jsp?dir=to&id=[pID]"
					  ;~ : (view = "rta") ? "CreateModifyRta.jsp?id=[pID]"
					  ;~ : (view = "m" || view = "mod" || view = "modify") ? "Modify.jsp?id=[pID]"
					  ;~ : (view = "p" || view = "print") ? "View_noMenu.jsp?id=[pID]&flowPic=false&printFriendly=True"
					  ;~ : "StructureManagement.jsp?id=[pID]")
	
		;~ ErrorLevel=
		;~ return Replace(viewURL, "[pID]", rID)
	;~ }







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
		
	Group: About
	
	Author: zzzooo10
	Link: http://www.autohotkey.com/forum/viewtopic.php?p=462061#462061

	Thanks to Tuncay for the idea: http://www.autohotkey.com/forum/viewtopic.php?t=74496

	Licence:
		Use in source, library and binary form is permitted.
		Redistribution and modification must meet the following condition:
		- My nickname (zzzooo10) and the origin (link) must be reproduced by binaries, or attached in the documentation.
		ALL MY SOFTWARE IS PROVIDED "AS IS" WITHOUT ANY EXPRESSED OR IMPLIED WARRANTIES.
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


















