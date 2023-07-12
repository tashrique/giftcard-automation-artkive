Dim appRef
Dim javaScriptFile
Dim fso, currentDirectory

Set fso = CreateObject("Scripting.FileSystemObject")
currentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)

Set appRef = CreateObject("Photoshop.Application")
javaScriptFile = currentDirectory & "\CSV to PSD - Giftcard Script.jsx"
appRef.DoJavaScriptFile(javaScriptFile)
