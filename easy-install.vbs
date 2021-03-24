' -------------------------------------------------------------------------------
' Beautify_Document_Addin installation script Ver.0.1
' -------------------------------------------------------------------------------
' References
' 1. How to install VBA addins via VBScript:
' https://www.aruse.net/entry/2018/09/13/081734#Excel-%E3%82%A2%E3%83%89%E3%82%A4%E3%83%B3%E3%81%AE%E3%82%A4%E3%83%B3%E3%82%B9%E3%83%88%E3%83%BC%E3%83%AB
' 修正
'
' 2. Relax-Tools-Addin
' https://github.com/RelaxTools/RelaxTools-Addin
' -------------------------------------------------------------------------------
Option Explicit
On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin
Dim strPath
Dim objWshShell
Dim objFileSys

'Add-in information
addInName = "Beautify Documents Addin"
addInFileName = "BeautifyAddin.xlam" 

IF MsgBox("Execute " & addInName & " installation?", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'Instantiate Excel
With CreateObject("Excel.Application") 

   'Intallation location
   '(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
   strPath = .UserLibraryPath

   IF Not objFileSys.FolderExists(strPath) THEN
        objFileSys.CreateFolder(strPath)
   END IF

   installPath = strPath & addInFileName

   'Copy file
   objFileSys.CopyFile  "bin\" & addInFileName ,installPath , True

   'Register Addin
   .Workbooks.Add
   Set objAddin = .AddIns.Add(installPath, True) 
   objAddin.Installed = True

   'Quit
   .Quit

End With

IF Err.Number = 0 THEN
   MsgBox "Completed installation", vbInformation
ELSE
   MsgBox Err.Description
End IF

Set objWshShell = Nothing
Set objFileSys = Nothing
