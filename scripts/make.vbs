' @author Hai Lu
' Read system.properties file, generate inno setup compiler script

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

Dim sKey
For Each sKey In dicProps
WScript.Echo sKey, "=", dicProps( sKey )
Next
Dim issName : issName = "inno-setup-script.iss"
' Create target dir if not exists
Dim targetDir : targetDir = ".\target"
If Not oFS.FolderExists(targetDir) Then
	oFS.CreateFolder (targetDir)
End If
Dim issSource
Set foIn = oFS.OpenTextFile(".\setup\" & issName)
Set foOut = oFS.CreateTextFile(targetDir & "\" & issName,True)
issSource = foIn.ReadAll

Dim curDir :curDir= oFS.GetAbsolutePathName(".")
Dim version : version = dicProps("app.version")
' Check buildnumer exists, append to app version
If oFS.FileExists(dicProps("build.number.file")) Then
   Dim rBN : Set rBN = oFS.OpenTextFile(dicProps("build.number.file"))
   Dim bn : bn = rBN.ReadAll
   version = version & "." & bn
End If

' loop all key and replace script content
Dim strKey
For Each strKey In dicProps.Keys()
	If StrComp(strKey, "license.file", vbTextCompare) = 0 Or  StrComp(strKey, "info.before.file", vbTextCompare) = 0 Or StrComp(strKey, "info.after.file", vbTextCompare) = 0 Or StrComp(strKey, "setup.icon.file", vbTextCompare) = 0 Then
		issSource = Replace(issSource, "%" & strKey & "%", curDir & "\" & dicProps(strKey))
	ElseIf StrComp(strKey, "app.version", vbTextCompare) = 0 Then
		issSource = Replace(issSource, "%" & strKey & "%", version)
	Else
		issSource = Replace(issSource, "%" & strKey & "%", dicProps(strKey))    
	End If
Next
issSource = Replace(issSource, "%output.dir%", curDir & "\target")  
issSource = Replace(issSource, "%source.file%", curDir & "\" & dicProps("app.db.file")) 
issSource = Replace(issSource, "%base.filename%", dicProps("app.file.name") & "-v" & version)

foOut.Write issSource
foIn.Close
foOut.Close

