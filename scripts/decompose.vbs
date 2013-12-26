' @author Oliver, Hai Lu

' Usage:
'  CScript decompose.vbs <input file> <path>

' Converts all modules, classes, forms and macros from an Access Project file (.adp) <input file> to
' text and saves the results in separate files to <path>.  Requires Microsoft Access.
'

Option Explicit

const acForm = 2
const acModule = 5
const acMacro = 4
const acReport = 3

' BEGIN CODE
WScript.Echo "Read project.properties ..."
Dim oFS : Set oFS = CreateObject( "Scripting.FileSystemObject" )
Dim sPFSpec : sPFSpec = ".\project.properties"
Dim dicProps : Set dicProps = CreateObject( "Scripting.Dictionary" )
Dim oTS : Set oTS = oFS.OpenTextFile( sPFSpec )
Dim sSect : sSect = ""
Do Until oTS.AtEndOfStream
Dim sLine : sLine = Trim( oTS.ReadLine )
If "" <> sLine Then
If "#" = Left( sLine, 1 ) Then
sSect = sLine
Else
If "" = sSect Then
Else
Dim aParts : aParts = Split( sLine, "=" )
If 1 <> UBound( aParts ) Then
Else
dicProps(Trim( aParts( 0 ) ) ) = Trim( aParts( 1 ) )
End If
End If
End If
End If
Loop
oTS.Close

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

dim sADPFilename
If (WScript.Arguments.Count = 0) then
    MsgBox "No parameter found!", vbExclamation, "Error"
    Wscript.Quit()
End if
sADPFilename = fso.GetAbsolutePathName(WScript.Arguments(0))

Dim sExportpath
If (WScript.Arguments.Count = 1) then
    sExportpath = ""
else
    sExportpath = WScript.Arguments(1)
End If


exportModulesTxt sADPFilename, sExportpath

If (Err <> 0) and (Err.Description <> NULL) Then
    MsgBox Err.Description, vbExclamation, "Error"
    Err.Clear
End If

Function exportModulesTxt(sADPFilename, sExportpath)
    Dim myComponent
    Dim sModuleType
    Dim sTempname
    Dim sOutstring

    dim myType, myName, myPath, sStubADPFilename
    myType = fso.GetExtensionName(sADPFilename)
    myName = fso.GetBaseName(sADPFilename)
    myPath = fso.GetParentFolderName(sADPFilename)

    If (sExportpath = "") then
        sExportpath = myPath & "\source\"
    End If
	If oFS.FolderExists(sExportpath) Then
		oFS.DeleteFolder(sExportpath & "*"),True
		oFS.DeleteFile(sExportpath & "*"),True
	End If
    sStubADPFilename = sExportpath & myName & "_stub." & myType

    WScript.Echo "copy stub to " & sStubADPFilename & "..."
    On Error Resume Next
        fso.CreateFolder(sExportpath)
    On Error Goto 0
    fso.CopyFile sADPFilename, sStubADPFilename

    WScript.Echo "starting Access..."
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")
    WScript.Echo "opening " & sStubADPFilename & " ..."
    If (Right(sStubADPFilename,4) = ".adp") Then
        oApplication.OpenAccessProject sStubADPFilename
    Else
        oApplication.OpenCurrentDatabase sStubADPFilename
    End If

    oApplication.Visible = false

    dim dctDelete
    Set dctDelete = CreateObject("Scripting.Dictionary")
    WScript.Echo "exporting..."
    Dim myObj
    For Each myObj In oApplication.CurrentProject.AllForms
        WScript.Echo "  " & myObj.fullname
        SaveAsText oApplication, acForm, myObj.fullname, sExportpath, myObj.fullname & ".form", dicProps
        oApplication.DoCmd.Close acForm, myObj.fullname
        dctDelete.Add "FO" & myObj.fullname, acForm
    Next
    For Each myObj In oApplication.CurrentProject.AllModules
        WScript.Echo "  " & myObj.fullname
        SaveAsText oApplication, acModule, myObj.fullname, sExportpath, myObj.fullname & ".bas", dicProps
        dctDelete.Add "MO" & myObj.fullname, acModule
    Next
    For Each myObj In oApplication.CurrentProject.AllMacros
        WScript.Echo "  " & myObj.fullname
        SaveAsText oApplication, acMacro, myObj.fullname, sExportpath, myObj.fullname & ".mac", dicProps
        dctDelete.Add "MA" & myObj.fullname, acMacro
    Next
    For Each myObj In oApplication.CurrentProject.AllReports
        WScript.Echo "  " & myObj.fullname
        SaveAsText oApplication, acReport, myObj.fullname, sExportpath, myObj.fullname & ".report", dicProps
        dctDelete.Add "RE" & myObj.fullname, acReport
    Next

    WScript.Echo "deleting..."
    dim sObjectname
    For Each sObjectname In dctDelete
        WScript.Echo "  " & Mid(sObjectname, 3)
        oApplication.DoCmd.DeleteObject dctDelete(sObjectname), Mid(sObjectname, 3)
    Next

    oApplication.CloseCurrentDatabase
    oApplication.CompactRepair sStubADPFilename, sStubADPFilename & "_"
    oApplication.Quit

    fso.CopyFile sStubADPFilename & "_", sStubADPFilename
    fso.DeleteFile sStubADPFilename & "_"


End Function

Function SaveAsText(oApplication, acObj, fullName, path, fileName, dicProps)
	Dim desPath, script
	Dim check : check = False
	If SaveModule(oApplication, acObj, fullName, path & "test", fileName, "src.test") Then
		check = True
	End If
	If SaveModule(oApplication, acObj, fullName, path & "test\lib", fileName, "src.test.lib") Then
		check = True
	End If
	If SaveModule(oApplication, acObj, fullName, path & "common", fileName, "src.common") Then
		check = True
	End If
	If Not check Then
		SaveFile oApplication, acObj, fullName, path & "main", fileName
	End If
End Function

Function SaveModule(oApplication, acObj, fullName, path, fileName, key)
	CheckDir(path)
	Dim source : source = dicProps(key)
	Dim check : check = False
	Dim str
	If Not StrComp(source,"",vbTextCompare) = 0 Then
		Dim list : list = Split(source, ",")
		For Each str In list
			If StrComp(Trim(str), fileName, vbTextCompare) = 0 Then
				SaveFile oApplication, acObj, fullName, path, fileName
				check = True
			End If
		Next
	End If
	SaveModule = check
End Function

Function SaveFile(oApplication, acObj, fullName, path, fileName)
	CheckDir path
	oApplication.SaveAsText acObj, fullName, path & "\" & fileName
End Function

Function CheckDir(path)
	Dim oFS : Set oFS = CreateObject("Scripting.FileSystemObject")
	If Not oFS.FolderExists(path) Then
		WScript.Echo "Create dir " & path
		oFS.CreateFolder (path)
	End If
End Function

Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.source & ":" & vbCrLf & _
               "    Description: " & Err.Description & vbCrLf & _
               "    Code: " & Err.Number & vbCrLf
    getErr = strError
End Function