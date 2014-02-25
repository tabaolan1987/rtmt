' @author Hai Lu
' Zip output setup file
If (WScript.Arguments.Count = 0) then
    MsgBox "No parameter found.!", vbExclamation, "Error"
    Wscript.Quit()
End if
Dim env
env = WScript.Arguments(0)

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

Dim version : version = dicProps("app.version")
' Check buildnumer exists, append to app version
If oFS.FileExists(dicProps("build.number.file")) Then
   Set oTS = oFS.OpenTextFile(dicProps("build.number.file"))
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
						dicProps(Trim( aParts( 0 ) ) ) = Trim( aParts( 1 ) )
					End If
				End If
			End If
		End If
	Loop
   version = version & "." & dicProps("build.number")
End If
Dim strNow, strDD, strMM, strYYYY
strYYYY = DatePart("yyyy",Now())
strMM = Right("0" & DatePart("m",Now()),2)
strDD = Right("0" & DatePart("d",Now()),2)
strNow = strDD & "-" & strMM & "-" & strYYYY
WScript.Echo "Get current date: " & strNow
Dim fromFile : fromFile = oFS.GetAbsolutePathName(".\target\" & dicProps("app.file.name") & "-v" & version & ".exe")
Dim zipFile : zipFile = oFS.GetAbsolutePathName(".\target\" & dicProps("app.file.name") & "-v" & version _
				& "_" & UCase(env) _
				& "_" & strNow & ".zip")
WScript.Echo "Create zip file: " & zipFile & " from file: " & fromFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.OpenTextFile(zipfile, 2, True).Write "PK" & Chr(5) & Chr(6) _
  & String(18, Chr(0))

Set ShellApp = CreateObject("Shell.Application")
Set zip = ShellApp.NameSpace(zipfile)
zip.CopyHere fromFile
WScript.Sleep 5000