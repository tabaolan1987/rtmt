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
        cRes = CurrentDb.Name
        nPos = Len(cRes)
        Do Until Right(cRes, 1) = "\"
            nPos = nPos - 1
            cRes = Left(cRes, nPos)
        Loop
        dbPath = cRes
    End If
    CurrentDbPath = dbPath
End Function

Function IsExist(path As String) As Boolean
    IsExist = Dir(path) <> ""
End Function

Function Delete(path As String) As Boolean
    If IsExist(path) Then
        SetAttr path, vbNormal
        Kill path
        Delete = True
    Else
        Delete = False
    End If
End Function

Public Function ReadSSFile(Name As String) As String()
    Dim path As String
    Dim arraySize As Integer
    Dim sInput As String
    Dim check As Boolean
    Dim i As Long
    Dim tmpList() As String
    Dim ln As String
    path = FileHelper.CurrentDbPath & Constants.SS_DIR & Name & ".ss"
    If IsExist(path) Then
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
    Dim Name As String
    Dim v As Variant
    If IsExist(desFilePath) Then
        Delete desFilePath
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
                            Logger.LogDebug "FileHelper.SaveAsCSV", "Sheet name: " & v.Name
                            If Not StringHelper.IsEqual(v.Name, WorkSheet, True) Then
                                check = True
                                'v.Delete
                            End If
                        Next v
                        If check Then
                            For Each v In .Sheets
                                'Logger.LogDebug "FileHelper.SaveAsCSV", "Sheet name: " & v.Name
                                If Not StringHelper.IsEqual(v.Name, WorkSheet, True) Then
                                
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
    Delete tmpSource
    Delete tmpStr
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

Public Function DuplicateAsTemporary(file As String) As String
    Dim desFile As String
    desFile = tmpDir & StringHelper.GetGUID
    Dim fso As New Scripting.FileSystemObject
    fso.CopyFile file, desFile, True
    Set fso = Nothing
    DuplicateAsTemporary = desFile
End Function

Public Function FileLastModified(strFullFileName As String)
    If IsExist(strFullFileName) Then
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