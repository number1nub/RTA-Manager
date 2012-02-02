Attribute VB_Name = "INIread"

Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long






' ___________________________________________________________________________________________________
' ===================================================================================================
'   Returns the value of a given INI key within the specified INI section
'
'   Parameters:
'       IniFile     - Full path to INI file
'       SectionName - INI Section to look within
'       KeyName     - Key whos value will be returned
'
'   Group: About
'       - *Written by:* Rameen Bakhtiary
'       - *Last modified:*  2012-01-11
' ___________________________________________________________________________________________________
' ===================================================================================================
Function INIread(ByVal filePath As String, ByVal SectionName As String, ByVal keyName As String)
        Dim RetVal As String, Exec As Integer, GetIni As String
'        FileExists filePath
        RetVal = String(255, 0)
        
        Exec = GetPrivateProfileString(SectionName, keyName, "", RetVal, Len(RetVal), filePath)
        
        If Len(Exec) = 0 Then
            INIread = ""
        Else
            INIread = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
        End If
End Function




' ___________________________________________________________________________________________________
' ===================================================================================================
' Sub: regRead
'
'   Reads a value from the Windows registry. Used to check the user's version.
'
' Last Modified: 2012-01-17
' ___________________________________________________________________________________________________
' ===================================================================================================
Sub cmdRead(keyvalue As String, KEY_QUERY_VALUE As String)
    Dim strValue As String * 256
    Dim lngRetval As Long
    Dim lngLength As Long
    Dim lngKey As Long

    If RegOpenKeyEx("HKEY_CURRENT_USER\Software\Halliburton RTA Manager", keyvalue, 0, KEY_QUERY_VALUE, lngKey) Then
    End If

    lngLength = 256

    'Retrieve the value of the key
    lngRetval = RegQueryValueEx(lngKey, keyvalue, 0, 0, ByVal strValue, lngLength)
    MsgBox Left(strValue, lngLength)

    'Close the key
    RegCloseKey (lngKey)
End Sub





Sub Read_registry_Value(keyName As String, value As String)
    Dim Shell As Object
    Dim keyvalue As String
    
    
    
    Set Shell = CreateObject("wscript.shell")
    On Error Resume Next
    keyvalue = Shell.regread(keyName & value)
    If Err.Number = 0 Then
        MsgBox "Invalid Registry Entry"
    Else
        MsgBox keyvalue
    End If
    On Error GoTo 0
End Sub















