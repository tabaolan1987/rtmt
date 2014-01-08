' @author Hai Lu
' 

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
MakeACCDESysCmd "D:\GIT\rolemapping-repo\rolemapping 2.accdb", "D:\GIT\rolemapping-repo\target\rolemapping.accde"

Public Function MakeACCDESysCmd(InPath, OutPath)
	Dim app 
	Set app = CreateObject("Access.Application")
	With app
        .AutomationSecurity = 1
        .UserControl = True
        .SysCmd 603, InPath, OutPath
        .Quit
    End With
End Function