option explicit

Const vbext_ct_ClassModule = 2
Const vbext_ct_Document = 100
Const vbext_ct_MSForm = 3
Const vbext_ct_StdModule = 1

Main

Sub Main
	Dim xl
	Dim fs
	Dim WBook
	Dim VBComp
	Dim Sfx
	Dim ExportFolder
	Dim classFolder
	Dim moduleFolder
	Dim formFolder
	Dim outPath
	
	
	
	'===================================================
	'		ENSURE AN EXCEL FILE WAS PASSED
	'===================================================
	If Wscript.Arguments.Count <> 1 Then 
		MsgBox "Invalid Excel file..."
		Exit Sub
	End If
	
	Set xl = CreateObject("Excel.Application")
	Set fs = CreateObject("Scripting.FileSystemObject")
	xl.Visible = false
	Set WBook = xl.Workbooks.Open(Trim(wScript.Arguments(0)))


	'===================================================
	'		CREATE FOLDERS FOR SOURCE FILES
	'===================================================
	ExportFolder = WBook.Path & "\" & fs.GetBaseName(WBook.Name) & " VBA Source"
	classFolder = ExportFolder & "\Class Modules"
	moduleFolder = ExportFolder & "\Modules"
	formFolder = ExportFolder & "\Forms"
	
	If Not fs.FolderExists(ExportFolder) Then fs.CreateFolder(ExportFolder)
	If Not fs.FolderExists(classFolder) Then fs.CreateFolder(classFolder)
	If Not fs.FolderExists(moduleFolder) Then fs.CreateFolder(moduleFolder)
	If Not fs.FolderExists(formFolder) Then fs.CreateFolder(formFolder)
	
	'===================================================
	'		PARSE THROUGH ALL VB OBJECTS & EXPORT THE 
	'		SOURCE FILES
	'===================================================
	For Each VBComp In WBook.VBProject.VBComponents
		
		Select Case VBComp.Type
			
			''~~~~~~~~~~~~~~~~~~~~~~
			'' Worksheet/Workbook
			''
			Case vbext_ct_Document
				Sfx = ".cls"
				outPath = ExportFolder & "\" & VBComp.Name & Sfx
			
			''~~~~~~~~~~~~~~~~
			'' Class Module
			''
			Case vbext_ct_ClassModule
				Sfx = ".cls"
				outPath = classFolder & "\" & VBComp.Name & Sfx
				
			''~~~~~~~~~
			''  Form
			''
			Case vbext_ct_MSForm
				Sfx = ".frm"
				outPath = formFolder & "\" & VBComp.Name & Sfx
				
			''~~~~~~~~~~~
			''  Module
			''
			Case vbext_ct_StdModule
				Sfx = ".bas"
				outPath = moduleFolder & "\" & VBComp.Name & Sfx
				
			''~~~~~~~~~~~~~~~~~~
			'' Something else	
			''		
			Case Else
				outPath = ""
		End Select
		
		
		If outPath <> "" Then
			On Error Resume Next
			Err.Clear
			
			''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'' Export the component if it was one of the above types
			''
			VBComp.Export outPath
			If Err.Number <> 0 Then MsgBox "Failed to export " & outPath
			On Error Goto 0
		End If
	
	Next

	'===================================================
	'			CLOSE THE EXCEL FILE
	'===================================================
	xl.DisplayAlerts() = False
	xl.Quit
End Sub
