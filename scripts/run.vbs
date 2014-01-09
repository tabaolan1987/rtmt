Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
dim sADPFilename
If (WScript.Arguments.Count = 0) then
    MsgBox "Please select database!", vbExclamation, "Error"
    Wscript.Quit()
End if

sADPFilename = fso.GetAbsolutePathName(WScript.Arguments(0))

Dim sMacroName
If (WScript.Arguments.Count = 1) then
    sMacroName = "VbaUnitMain"
else
    sMacroName = WScript.Arguments(1)
End If

Dim oAccess
'Start Access and open the database.
set oAccess = CreateObject("Access.Application")
oAccess.Visible = False
'You will need to put the path to your own database here.
oAccess.OpenCurrentDatabase(sADPFilename)
'Run the macro
oAccess.Run sMacroName
'Quit Access without saving the database.
set oAccess = Nothing