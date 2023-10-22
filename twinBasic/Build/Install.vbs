const AddInName = "ACLibAddInStarter"
const AddInFileName = "AClibAddInStarter_win32.dll"
const MsgBoxTitle = "Install ACLib Add-In Starter"
const RibbonFileName = "ACLibAddInStarterRibbon.xml"

Dim AddInFileInstalled, CompletedMsg

MsgBox "Before updating the add-in, the add-in must not be loaded!" & chr(13) & _
       "For safety, close all Access instances.", , MsgBoxTitle & ": Information"

AddInFileInstalled = CopyFileAndRegisterAddIn(GetSourceFileFullName, GetDestFileFullName)
If AddInFileInstalled Then
	CompletedMsg = "Add-In installed in '" + GetAddInLocation + "'."
Else
	CompletedMsg = "Error! Add-In not installed."
End If

CopyRibbonXmlIfNotExists

MsgBox CompletedMsg, , MsgBoxTitle


'##################################################
' Functions

Function GetSourceFileFullName()
   GetSourceFileFullName = GetScriptLocation & AddInFileName 
End Function

Function GetDestFileFullName()
   GetDestFileFullName = GetAddInLocation & AddInFileName 
End Function

Function GetScriptLocation()
   With WScript
      GetScriptLocation = Replace(.ScriptFullName & ":", .ScriptName & ":", "") 
   End With
End Function

Function GetAddInLocation()
   GetAddInLocation = GetAppDataLocation & "Microsoft\AddIns\"
End Function

Function GetAppDataLocation()
   Set wsShell = CreateObject("WScript.Shell")
   GetAppDataLocation = wsShell.ExpandEnvironmentStrings("%APPDATA%") & "\"
End Function

Function FileCopy(SourceFilePath, DestFilePath)
   set fso = CreateObject("Scripting.FileSystemObject") 
   fso.CopyFile SourceFilePath, DestFilePath
   FileCopy = True
End Function

Function CopyRibbonXmlIfNotExists()
   ribbonFilePath = GetAddInLocation() & RibbonFileName
   Set fso = CreateObject("Scripting.FileSystemObject")
   if Not fso.FileExists(ribbonFilePath) then
      FileCopy GetScriptLocation() & RibbonFileName, GetAddInLocation() & RibbonFileName
   end if
End Function

Function DeleteAddInFiles()
   Set fso = CreateObject("Scripting.FileSystemObject")
   DeleteFile fso, GetDestFileFullName()
End Function

Function DeleteFile(fso, File2Delete)
   if fso.FileExists(File2Delete) then
      fso.DeleteFile File2Delete
   end if
End Function

Function CopyFileAndRegisterAddIn(SourceFilePath, DestFilePath)

   IF Not FileCopy(SourceFilePath, DestFilePath) Then
      Exit Function
   End If

   With CreateObject("WScript.Shell")
       .Exec "regsvr32 /s """ & DestFilePath & """"
   End With
   
   CopyFileAndRegisterAddIn = True
 
End Function
