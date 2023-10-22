const AddInName = "ACLibAddInStarter"
const AddInFileName = "AClibAddInStarter_win32.dll"
const MsgBoxTitle = "Uninstall ACLib Add-In Starter"

Dim AddInFileUninstalled, CompletedMsg

MsgBox "Before updating the add-in, the add-in must not be loaded!" & chr(13) & _
       "For safety, close all Access instances.", , MsgBoxTitle & ": Information"

AddInFileUninstalled = UnRegisterAddIn(GetDestFileFullName)
If AddInFileUninstalled Then
    CompletedMsg = "Add-In uninstalled."
    DeleteAddInFiles
Else
    CompletedMsg = "Error! Add-In not uninstalled."
End If

MsgBox CompletedMsg, , MsgBoxTitle


'##################################################
' Functions

Function GetDestFileFullName()
   GetDestFileFullName = GetAddInLocation & AddInFileName 
End Function

Function GetAddInLocation()
   GetAddInLocation = GetAppDataLocation & "Microsoft\AddIns\"
End Function

Function GetAppDataLocation()
   Set wsShell = CreateObject("WScript.Shell")
   GetAppDataLocation = wsShell.ExpandEnvironmentStrings("%APPDATA%") & "\"
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

Function UnRegisterAddIn(DestFilePath)

   With CreateObject("WScript.Shell")
       .Exec "regsvr32 /u /s """ & DestFilePath & """"
   End With
   
   UnRegisterAddIn = True
 
End Function
