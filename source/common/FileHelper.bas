'@author Hai Lu
' Work with File & Dir
Option Explicit
Const ForReading = 1

Private dbPath As String
Private tmpDirPath As String

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If


Public Function WaitForFileClose(FileName As String, ByVal TestIntervalMilliseconds As Double, _
    ByVal TimeOutMilliseconds As Double) As Boolean

Dim StartTickCount As Double
Dim EndTickCount As Double
Dim TickCountNow As Double
Dim FileIsOpen As Boolean
Dim done As Boolean
Dim CancelKeyState As Double

FileIsOpen = IsFileOpen(FileName:=FileName)
If FileIsOpen = False Then
    WaitForFileClose = True
    Exit Function
End If

If TestIntervalMilliseconds <= 0 Then
    TestIntervalMilliseconds = 500
End If

'CancelKeyState = Application.EnableCancelKey
'Application.EnableCancelKey = xlErrorHandler
On Error GoTo ErrHandler:
StartTickCount = GetTickCount()
If TimeOutMilliseconds <= 0 Then
    EndTickCount = -1
Else
    EndTickCount = StartTickCount + TimeOutMilliseconds
End If

done = False
Do Until done
    If IsFileOpen(FileName:=FileName) = False Then
        WaitForFileClose = True
   '     Application.EnableCancelKey = CancelKeyState
        Exit Function
    End If
    Sleep dwMilliseconds:=TestIntervalMilliseconds
    TickCountNow = GetTickCount()
    If EndTickCount > 0 Then
        If TickCountNow >= EndTickCount Then
            WaitForFileClose = Not (IsFileOpen(FileName))
  '          Application.EnableCancelKey = CancelKeyState
            Exit Function
        Else
        End If
    Else
        If IsFileOpen(FileName:=FileName) = False Then
            WaitForFileClose = True
 '           Application.EnableCancelKey = CancelKeyState
            Exit Function
        End If
        
    End If
    DoEvents
Loop
Exit Function

ErrHandler:
'Application.EnableCancelKey = CancelKeyState
WaitForFileClose = False

End Function


Private Function IsFileOpen(FileName As String) As Boolean
Dim FileNum As Integer
Dim ErrNum As Integer

On Error Resume Next
If FileName = vbNullString Then
    IsFileOpen = False
    Exit Function
End If
If Dir(FileName) = vbNullString Then
    IsFileOpen = False
    Exit Function
End If
FileNum = FreeFile()
Err.Clear
Open FileName For Input Lock Read As #FileNum
ErrNum = Err.Number
On Error GoTo 0
Close #FileNum
Select Case ErrNum
    Case 0
        IsFileOpen = False
    Case 70
        IsFileOpen = True
    Case Else
        IsFileOpen = True
        
End Select
End Function

Function GetCSVFile() As String
    Dim fDialog As Object
    Set fDialog = Application.FileDialog(3)
    With fDialog
        .Title = "Select the CSV file to import"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Workbook", "*.xlsx", 1
        .Filters.Add "Excel 97-2003 Workbook", "*.xls", 2
        .Filters.Add "CSV Files", "*.csv", 3
        .Filters.Add "All Files", "*.*", 4
        If .Show = True Then
            GetCSVFile = .SelectedItems(1)
        End If
    End With
    Set fDialog = Nothing
End Function

Public Sub CheckDir(strDirPath As String)
    strDirPath = Replace(strDirPath, "/", "\")
    Dim listStr() As String, strTemp As String, str As String
    listStr = Split(strDirPath, "\")
    strTemp = ""
    Dim i As Integer
    For i = LBound(listStr) To UBound(listStr)
       strTemp = strTemp & listStr(i) & "\"
        If Dir(strTemp, vbDirectory) = "" Then
            MkDir strTemp
        End If
    Next i
End Sub

Public Function ReadQuery(stName As String, Optional qType As Integer)
    Logger.LogDebug "FileHelper.ReadQuery", "qType: " & CStr(qType)
    Dim queryPath As String: queryPath = ""
    queryPath = queryPath & Constants.QUERIES_DIR
    Select Case qType
        Case Constants.Q_CREATE: queryPath = queryPath & "create_table_" & stName
        Case Constants.Q_INSERT: queryPath = queryPath & "insert_" & stName
        Case Constants.Q_UPDATE: queryPath = queryPath & "update_" & stName
        Case Constants.Q_DELETE_ALL: queryPath = queryPath & "delete_all_" & stName
        Case Constants.Q_SELECT: queryPath = queryPath & "select_" & stName
        Case Constants.Q_CUSTOM: queryPath = queryPath & stName
    End Select
    queryPath = queryPath & ".sql"
    Logger.LogDebug "FileHelper.ReadQuery", "Path: " & queryPath
    ReadQuery = ReadFile(queryPath)
End Function

Public Function ReadFileFullPath(strFileFullPath As String) As String
    Dim nSourceFile As Integer, sText As String
   Close
   ''Get the number of the next free text file
   nSourceFile = FreeFile
   ''Write the entire file to sText
   Open strFileFullPath For Input As #nSourceFile
        sText = Input$(LOF(1), 1)
   Close
   ReadFileFullPath = sText
End Function

Public Function ReadFile(strFilePath As String) As String
   ReadFile = ReadFileFullPath(CurrentDbPath & strFilePath)
   
End Function

Function CurrentDbPath() As String
    If Len(dbPath) = 0 Then
        Dim cRes As String
        Dim nPos As Long
        cRes = CurrentDb.name
        nPos = Len(cRes)
        Do Until Right(cRes, 1) = "\"
            nPos = nPos - 1
            cRes = Left(cRes, nPos)
        Loop
        dbPath = cRes
    End If
    CurrentDbPath = dbPath
End Function

Function IsExistFile(path As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    IsExistFile = fso.FileExists(path)
    'IsExist = Dir(path) <> ""
End Function

Function IsExistFolder(path As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    IsExistFolder = fso.FolderExists(path)
    'IsExist = Dir(path) <> ""
End Function

Function DeleteFile(path As String) As Boolean
    If IsExistFile(path) Then
        SetAttr path, vbNormal
        Kill path
        DeleteFile = True
    Else
        DeleteFile = False
    End If
End Function

Function DeleteFolder(path As String) As Boolean
    If IsExistFolder(path) Then
        SetAttr path, vbNormal
        Kill path
        DeleteFolder = True
    Else
        DeleteFolder = False
    End If
End Function

Public Function ReadSSFile(name As String) As String()
    Dim path As String
    Dim arraySize As Integer
    Dim sInput As String
    Dim check As Boolean
    Dim i As Long
    Dim tmpList() As String
    Dim ln As String
    path = FileHelper.CurrentDbPath & Constants.SS_DIR & name & ".ss"
    If IsExistFile(path) Then
        Dim fso As Object
        Dim ReadFile As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ReadFile = fso.OpenTextFile(path, ForReading, False)
        Do Until ReadFile.AtEndOfStream = True
            ln = Trim(ReadFile.ReadLine)
            If StringHelper.StartsWith(ln, "#", True) = False And Len(ln) <> 0 Then
                ReDim Preserve tmpList(arraySize)
                tmpList(arraySize) = ln
                arraySize = arraySize + 1
            End If
        Loop
        ReadFile.Close
        Set fso = Nothing
        Set ReadFile = Nothing
    End If
    
    ReadSSFile = tmpList
End Function

Public Function SaveAsCSV(filePath As String, desFilePath As String, Optional WorkSheet As String) As Boolean
    Dim oExcel As New Excel.Application
    Dim i As Integer
    Dim WB As New Excel.Workbook
    Dim WS As Excel.Sheets
    Dim name As String
    Dim v As Variant
    If IsExistFile(desFilePath) Then
        DeleteFile desFilePath
    End If
    Dim check As Boolean
    check = False
    With oExcel
        .Visible = False
        .DisplayAlerts = False
                    Set WB = .Workbooks.Add(filePath)
                    ' Remove unused sheets
                    Logger.LogDebug "FileHelper.SaveAsCSV", "Sheet count: " & .Sheets.Count
                    If .Sheets.Count > 1 And Len(WorkSheet) <> 0 Then
                        For Each v In .Sheets
                            Logger.LogDebug "FileHelper.SaveAsCSV", "Sheet name: " & v.name
                            If Not StringHelper.IsEqual(v.name, WorkSheet, True) Then
                                check = True
                                'v.Delete
                            End If
                        Next v
                        If check Then
                            For Each v In .Sheets
                                'Logger.LogDebug "FileHelper.SaveAsCSV", "Sheet name: " & v.Name
                                If Not StringHelper.IsEqual(v.name, WorkSheet, True) Then
                                
                                    v.Delete
                                End If
                            Next v
                        End If
                    End If
                  
                        WB.SaveAs desFilePath, FileFormat:=6 ' Save as CSV
                   
                    WB.Close False
        .Quit
    End With
    Set oExcel = Nothing
    SaveAsCSV = check
End Function

Public Function TrimSourceFile(fileToRead As String, fileToWrite As String, LineToRemove() As Integer)
    Dim fso As Object
    Dim ReadFile As Object
    Dim writeFile As Object
    Dim repLine As Variant
    Dim ln As String
    Dim l As Long
    Dim tmpList() As String
    Dim arraySize As Integer
    Dim check As Boolean
    Dim rowCheck As Boolean
    Dim tmpCheck() As String
    Dim i As Integer
    Dim ltm As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ReadFile = fso.OpenTextFile(fileToRead, ForReading, False)
    Set writeFile = fso.CreateTextFile(fileToWrite, True, False)
    
    '# iterate the array and do the replacement line by line
    Do Until ReadFile.AtEndOfStream = True
        ln = ReadFile.ReadLine
        check = False
        For Each ltm In LineToRemove
            If (ltm = l + 1) Then
                Logger.LogDebug "FileHelper.TrimSourceFile", "Line to remove " & CStr(ltm)
                check = True
            End If
        Next
        If check = False Then
            If Len(ln) <> 0 Then
                
                rowCheck = False
                tmpCheck = Split(ln, ",")
                For i = LBound(tmpCheck) To UBound(tmpCheck)
                    If Len(Trim(tmpCheck(i))) <> 0 Then
                        rowCheck = True
                    End If
                Next i
                If rowCheck Then
                    'Logger.LogDebug "FileHelper.TrimSourceFile", "Readline " & CStr(l) & " . Text: " & ln
                    ReDim Preserve tmpList(arraySize)
                    tmpList(arraySize) = ln
                    arraySize = arraySize + 1
                End If
            End If
        End If
        l = l + 1
    Loop
    ReadFile.Close
    '# Write to the array items to the file
    writeFile.Write Join(tmpList, vbNewLine)
    writeFile.Close
    
    '# clean up
    Set ReadFile = Nothing
    Set writeFile = Nothing
    Set fso = Nothing
End Function

Public Function PrepareUserData(filePath As String, ss As SystemSetting) As String
    Dim tmpStr As String
    Dim tmpSource As String
    Dim outputCsv As String
    Dim fso As New Scripting.FileSystemObject
    If StringHelper.EndsWith(filePath, ".xlsx", True) Or _
        StringHelper.EndsWith(filePath, ".xls", True) Or _
        StringHelper.EndsWith(filePath, ".csv", True) Then
        tmpSource = tmpDir & StringHelper.GetGUID
        Logger.LogDebug "FileHelper.PrepareUserData", "Copy file " & filePath & " to " & tmpSource
        fso.CopyFile filePath, tmpSource, True
        
        tmpStr = tmpDir & StringHelper.GetGUID & ".csv"
        Logger.LogDebug "FileHelper.PrepareUserData", "Convert file " & tmpSource & " to CSV file " & tmpStr
        FileHelper.SaveAsCSV tmpSource, tmpStr, ss.WorkSheet
            outputCsv = tmpDir & StringHelper.GetGUID & ".csv"
            Logger.LogDebug "FileHelper.PrepareUserData", "Trim unused rows " & tmpStr & " to CSV file " & outputCsv
            TrimSourceFile tmpStr, outputCsv, ss.LineToRemove
            PrepareUserData = outputCsv
        Set fso = Nothing
        DeleteFile tmpSource
        DeleteFile tmpStr
    Else
        PrepareUserData = ""
    End If
End Function

Public Function tmpDir() As String
    If Len(tmpDirPath) = 0 Then
        Dim fso As New Scripting.FileSystemObject
        tmpDirPath = fso.GetSpecialFolder(TemporaryFolder).path
        If Not StringHelper.EndsWith(tmpDirPath, "\", True) Then
            tmpDirPath = tmpDirPath & "\"
        End If
        tmpDirPath = tmpDirPath & "rmt\"
        CheckDir (tmpDirPath)
        Set fso = Nothing
    End If
    tmpDir = tmpDirPath
End Function

Public Function MoveFile(mFrom As String, mTo As String)
    Dim fso As New Scripting.FileSystemObject
    fso.MoveFile mFrom, mTo
    Set fso = Nothing
End Function

Public Function CopyFile(mFrom As String, mTo As String)
    Dim fso As New Scripting.FileSystemObject
    fso.CopyFile mFrom, mTo, True
    Set fso = Nothing
End Function

Public Function DuplicateAsTemporary(file As String) As String
    Dim desFile As String
    desFile = tmpDir & StringHelper.GetGUID
    Dim fso As New Scripting.FileSystemObject
    fso.CopyFile file, desFile, True
    Set fso = Nothing
    DuplicateAsTemporary = desFile
End Function

Public Function FileLastModified(strFullFileName As String)
    If IsExistFile(strFullFileName) Then
        Dim fs As New Scripting.FileSystemObject, f As Object, s As String
        Set f = fs.GetFile(strFullFileName)
        s = UCase(strFullFileName) & vbCrLf
        FileLastModified = f.DateLastModified
        Set fs = Nothing: Set f = Nothing
    Else
        FileLastModified = ""
    End If
End Function

Public Function ImageLoaderSource() As String
    ImageLoaderSource = "=""" & FileHelper.CurrentDbPath & "data\images\loader.html"""
End Function



Public Sub Zip( _
    ZipFile As String, _
    InputFile As String _
)
On Error GoTo ErrHandler
    Dim fso As Object 'Scripting.FileSystemObject
    Dim oApp As Object 'Shell32.Shell
    Dim oFld As Object 'Shell32.Folder
    Dim oShl As Object 'WScript.Shell
    Dim i As Long
    Dim l As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(ZipFile) Then
        'Create empty ZIP file
        fso.CreateTextFile(ZipFile, True).Write _
            "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
    End If

    Set oApp = CreateObject("Shell.Application")
    Set oFld = oApp.Namespace(CVar(ZipFile))
    i = oFld.Items.Count
    oFld.CopyHere (InputFile)

    Set oShl = CreateObject("WScript.Shell")

    'Search for a Compressing dialog
    Do While oShl.AppActivate("Compressing...") = False
        If oFld.Items.Count > i Then
            'There's a file in the zip file now, but
            'compressing may not be done just yet
            Exit Do
        End If
        If l > 30 Then
            '3 seconds has elapsed and no Compressing dialog
            'The zip may have completed too quickly so exiting
            Exit Do
        End If
        DoEvents
        Sleep 100
        l = l + 1
    Loop

    ' Wait for compression to complete before exiting
    Do While oShl.AppActivate("Compressing...") = True
        DoEvents
        Sleep 100
    Loop

ExitProc:
    On Error Resume Next
        Set fso = Nothing
        Set oFld = Nothing
        Set oApp = Nothing
        Set oShl = Nothing
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case Else
            MsgBox "Error " & Err.Number & _
                   ": " & Err.Description, _
                   vbCritical, "Unexpected error"
    End Select
    Resume ExitProc
    Resume
End Sub

Public Sub UnZip( _
    ZipFile As String, _
    Optional TargetFolderPath As String = vbNullString, _
    Optional OverwriteFile As Boolean = False _
)
On Error GoTo ErrHandler
    Dim oApp As Object
    Dim fso As Object
    Dim fil As Object
    Dim DefPath As String
    Dim strDate As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(TargetFolderPath) = 0 Then
        DefPath = CurrentProject.path & ""
    Else
        If fso.FolderExists(TargetFolderPath) Then
            DefPath = TargetFolderPath & ""
        Else
            Err.Raise 53, , "Folder not found"
        End If
    End If

    If fso.FileExists(ZipFile) = False Then
        MsgBox "System could not find " & ZipFile _
            & " upgrade cancelled.", _
            vbInformation, "Error Unziping File"
        Exit Sub
    Else
        'Extract the files into the newly created folder
        Set oApp = CreateObject("Shell.Application")

        With oApp.Namespace(ZipFile & "")
            If OverwriteFile Then
                For Each fil In .Items
                    If fso.FileExists(DefPath & fil.name) Then
                        Kill DefPath & fil.name
                    End If
                Next
            End If
            oApp.Namespace(CVar(DefPath)).CopyHere .Items
        End With

        On Error Resume Next
        Kill Environ("Temp") & "Temporary Directory*"

        'Kill zip file
        Kill ZipFile
    End If

ExitProc:
    On Error Resume Next
    Set oApp = Nothing
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected error"
    End Select
    Resume ExitProc
    Resume
End Sub

Function FileSaveDialog(InitialFileName As String)
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    fd.InitialFileName = InitialFileName
    fd.AllowMultiSelect = False
    If CBool(fd.Show) Then
            FileSaveDialog = fd.SelectedItems(fd.SelectedItems.Count)
    Else
    
    End If
End Function