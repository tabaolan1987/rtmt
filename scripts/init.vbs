' @author Hai Lu
' Read environment *.properties file, generate template file frol .template folder to .data folder
If (WScript.Arguments.Count = 0) then
    MsgBox "No parameter found.!", vbExclamation, "Error"
    Wscript.Quit()
End if
Dim env
env = WScript.Arguments(0)

Dim sLine
Dim aParts
Dim oFS : Set oFS = CreateObject( "Scripting.FileSystemObject" )
Dim sPFSpec : sPFSpec = ".\env\" & env & ".properties"
WScript.Echo "Read " & sPFSpec
Dim envProps : Set envProps = CreateObject( "Scripting.Dictionary" )
Dim oTS : Set oTS = oFS.OpenTextFile( sPFSpec )
Dim sSect : sSect = ""
Do Until oTS.AtEndOfStream
	sLine = Trim( oTS.ReadLine )
	If "" <> sLine Then
		If "#" = Left( sLine, 1 ) Then
			sSect = sLine
		Else
			If "" = sSect Then
				WScript.Echo sLine
			Else
				
				aParts = Split( sLine, "=" )
				If 1 <> UBound( aParts ) Then
				Else
					envProps(Trim( aParts( 0 ) ) ) = Trim( aParts( 1 ) )
				End If
			End If
		End If
	End If
Loop
oTS.Close

WScript.Echo "Read project.properties ..."
Dim prjProps : Set prjProps = CreateObject( "Scripting.Dictionary" )
Set oTS = oFS.OpenTextFile( ".\project.properties" )
sSect = ""
Do Until oTS.AtEndOfStream
	sLine = Trim( oTS.ReadLine )
	If "" <> sLine Then
		If "#" = Left( sLine, 1 ) Then
			sSect = sLine
		Else
			If "" = sSect Then
			Else
				aParts = Split( sLine, "=" )
				If 1 <> UBound( aParts ) Then
				Else
					prjProps(Trim( aParts( 0 ) ) ) = Trim( aParts( 1 ) )
				End If
			End If
		End If
	End If
Loop
oTS.Close
Dim curDir :curDir = oFS.GetAbsolutePathName(".")
Dim strKey
Dim fileSource
Dim fileName
Dim targetDir

For Each strKey In prjProps.Keys()
	envProps(strKey) = prjProps(strKey)
Next

Dim version : version = prjProps("app.version")
' Check buildnumer exists, append to app version
If oFS.FileExists(prjProps("build.number.file")) Then
   Dim rBN : Set rBN = oFS.OpenTextFile(prjProps("build.number.file"))
   Dim bn : bn = rBN.ReadLine
   version = version & "." & Trim(bn)
End If
envProps("env") = env
envProps("project.version") = version

For Each strKey In envProps.Keys()
	WScript.Echo strKey & " = " & envProps(strKey)
Next
CheckFolder ".\template", oFS, envProps

Public Function CheckFolder(folder, oFS, envProps)
	Dim f
	Dim fileName
	Dim targetDir
	Dim folderPath
	Dim foIn
	Dim foOut
	folderPath = oFS.GetAbsolutePathName(folder)
	targetDir = oFS.GetAbsolutePathName(".\data") &	Right(folderPath, Len(folderPath) - Len(oFS.GetAbsolutePathName(".\template")))
	If Not oFS.FolderExists(targetDir) Then
		WScript.Echo "Create new folder: " & targetDir
		oFS.CreateFolder (targetDir)
	End If
	
    For Each f In oFS.GetFolder(folderPath).Files
		fileName = oFS.GetAbsolutePathName(".\data") & Right(f.Path, Len(f.Path) - Len(oFS.GetAbsolutePathName(".\template")))
		WScript.Echo "Generate: " & fileName & ". From: " & f.Path
		Set foIn = oFS.OpenTextFile(f.Path)
		Set foOut = oFS.CreateTextFile(fileName, True)
		fileSource = foIn.ReadAll
		' loop all key and replace script content

		For Each strKey In envProps.Keys()
			fileSource = Replace(fileSource, "{%" & strKey & "%}", envProps(strKey))  
		Next
		foOut.Write fileSource
		foIn.Close
		foOut.Close
	Next
	For Each f In oFS.GetFolder(folderPath).SubFolders
		CheckFolder f.Path, oFS, envProps
	Next
End Function
