Attribute VB_Name = "Module1"
Sub Save_RTAload()
'
' Save_RTAload Macro
'

'
    Application.ScreenUpdating = False
    Sheets("RTAimport").Visible = True
    Sheets("RTAimport").Select
    Sheets("RTAimport").Copy
    
    
    ActiveWorkbook.SaveAs Filename:="C:\Users\hb52875\Documents\rtaLoad.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
   ChDir "\\corp.halliburton.com\team\WD\Business Development and Technology\General\Engineering Public\RTA Management Sheet"
        
        fName = "\\corp.halliburton.com\team\WD\Business Development and Technology\General\Engineering Public\RTA Management Sheet\" & _
            Format(Now(), "yyyy-m-d  hhmm ") & "(" & UserNameWindows & ")  " & UCase("cFilt") & ".xlsx"
    
        ActiveWorkbook.SaveAs Filename:=fName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
        
    ActiveWindow.Close
    
    Sheets("RTAimport").Select
    ActiveWindow.SelectedSheets.Visible = False
    Range("A1:B1").Select
    Application.ScreenUpdating = False
End Sub


Sub test()
    
    Sheets("Settings").Visible = xlSheetVisible
        Sheets("settings").Select
        Application.Goto ("cFilt")
        laboff = Selection(1, 1)
        Sheets("Settings").Visible = xlSheetHidden
    
    MsgBox laboff
End Sub
