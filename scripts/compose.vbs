' @author Oliver, Hai Lu

' Converts all modules, classes, forms and macros in a directory created by "decompose.vbs"
' and composes then into an Access Project file (.adp). This overwrites any existing Modules with the
' same names without warning!!!
' Requires Microsoft Access.

Option Explicit

const acForm = 2
const acModule = 5
const acMacro = 4
const acReport = 3

Const acCmdCompileAndSaveAllModules = &H7E

' BEGIN CODE
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

dim sADPFilename
If (WScript.Arguments.Count = 0) then
    MsgBox "No parameter found.!", vbExclamation, "Error"
    Wscript.Quit()
End if
sADPFilename = fso.GetAbsolutePathName(WScript.Arguments(0))

Dim sPath
If (WScript.Arguments.Count = 1) then
    sPath = ""
else
    sPath = WScript.Arguments(1)
End If

Dim prjEnv
If (WScript.Arguments.Count = 3) then
	prjEnv = WScript.Arguments(2)
else
    prjEnv = "dev"
End If
importModulesTxt sADPFilename, sPath

If (Err <> 0) and (Err.Description <> NULL) Then
    MsgBox Err.Description, vbExclamation, "Error"
    Err.Clear
End If

Function importModulesTxt(sADPFilename, sImportpath)
    Dim myComponent
    Dim sModuleType
    Dim sTempname
    Dim sOutstring

    ' Build file and pathnames
    dim myType, myName, myPath, sStubADPFilename
    myType = fso.GetExtensionName(sADPFilename)
    myName = fso.GetBaseName(sADPFilename)
    myPath = fso.GetParentFolderName(sADPFilename)

    ' if no path was given as argument, use a relative directory
    If (sImportpath = "") then
        sImportpath = myPath & "\source\"
    End If
    sStubADPFilename = sImportpath & myName & "_stub." & myType

    ' check for existing file and ask to overwrite with the stub
    if (fso.FileExists(sADPFilename)) Then
		If StrComp(prjEnv, "dev", vbTextCompare) = 0 Then
			WScript.StdOut.Write sADPFilename & " is existed. Override it? (y/n) "
			dim sInput
			sInput = WScript.StdIn.Read(1)
			if (sInput <> "y") Then
				WScript.Quit
			end if
		End If
        fso.CopyFile sADPFilename, sADPFilename & ".bak"
    end if

    fso.CopyFile sStubADPFilename, sADPFilename

    ' launch MSAccess
    WScript.Echo "starting Access..."
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")
    WScript.Echo "opening " & sADPFilename & " ..."
    If (Right(sStubADPFilename,4) = ".adp") Then
        oApplication.OpenAccessProject sADPFilename
    Else
        oApplication.OpenCurrentDatabase sADPFilename
    End If
    oApplication.Visible = false
	LoadModule oApplication, sImportpath & "main"
	LoadModule oApplication, sImportpath & "common"
	If StrComp(prjEnv, "dev", vbTextCompare) = 0 or StrComp(prjEnv, "test", vbTextCompare) = 0 Then
		LoadModule oApplication, sImportpath & "test"
		LoadModule oApplication, sImportpath & "test\lib"
	End If
	
    oApplication.RunCommand acCmdCompileAndSaveAllModules
    oApplication.Quit
End Function

Function LoadModule(oApplication, path)
	Dim folder
    Set folder = fso.GetFolder(path)
    ' load each file from the import path into the stub
    Dim myFile, objectname, objecttype
    for each myFile in folder.Files
        objecttype = fso.GetExtensionName(myFile.Name)
        objectname = fso.GetBaseName(myFile.Name)
        WScript.Echo "  " & objectname & " (" & objecttype & ")"

        if (objecttype = "form") then
            oApplication.LoadFromText acForm, objectname, myFile.Path
        elseif (objecttype = "bas") then
            oApplication.LoadFromText acModule, objectname, myFile.Path
        elseif (objecttype = "mac") then
            oApplication.LoadFromText acMacro, objectname, myFile.Path
        elseif (objecttype = "report") then
            oApplication.LoadFromText acReport, objectname, myFile.Path
        end if
    next
End Function
Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.source & ":" & vbCrLf & _
               "    Description: " & Err.Description & vbCrLf & _
               "    Code: " & Err.Number & vbCrLf
    getErr = strError
End Function