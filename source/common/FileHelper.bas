'@author Hai Lu
' Work with File & Dir
Option Explicit
Const ForReading = 1

Private dbPath As String

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If


Public Function WaitForFileClose(FileName As String, ByVal TestIntervalMilliseconds As Long, _
    ByVal TimeOutMilliseconds As Long) As Boolean

Dim StartTickCount As Long
Dim EndTickCount As Long
Dim TickCountNow As Long
Dim FileIsOpen As Boolean
Dim Done As Boolean
Dim CancelKeyState As Long

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

Done = False
Do Until Done
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
        .Filters.Add "CSV Files", "*.csv", 1
        .Filters.Add "Excel Workbook", "*.xlsx", 2
        .Filters.Add "Excel 97-2003 Workbook", "*.xls", 3
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
        Case Else: queryPath = queryPath & stName
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

Public Function ReadSSFile(name As String) As String()
    Dim path As String
    Dim arraySize As Integer
    Dim sInput As String
    Dim check As Boolean
    Dim i As Long
    Dim tmpList() As String
    Dim ln As String
    path = FileHelper.CurrentDbPath & Constants.SS_DIR & name & ".ss"
    If IsExist(path) Then
        Dim FSO As Object
        Dim ReadFile As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set ReadFile = FSO.OpenTextFile(path, ForReading, False)
        Do Until ReadFile.AtEndOfStream = True
            ln = Trim(ReadFile.ReadLine)
            If StringHelper.StartsWith(ln, "#", True) = False And Len(ln) <> 0 Then
                ReDim Preserve tmpList(arraySize)
                tmpList(arraySize) = ln
                arraySize = arraySize + 1
            End If
        Loop
        ReadFile.Close
        Set FSO = Nothing
        Set ReadFile = Nothing
    End If
    
    ReadSSFile = tmpList
End Function

Public Function SaveAsCSV(filePath As String, desFilePath As String)
    Dim oExcel As New Excel.Application
    Dim WB As New Excel.Workbook
    If IsExist(desFilePath) Then
        Delete desFilePath
    End If
    With oExcel
        .Visible = False
                    Set WB = .Workbooks.Add(filePath)
                    WB.SaveAs desFilePath, FileFormat:=6 ' Save as CSV
                    WB.Close False
        .Quit
    End With
    Set oExcel = Nothing
End Function

Public Function TrimSourceFile(fileToRead As String, fileToWrite As String, LineToRemove() As Integer)
    Dim FSO As Object
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
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ReadFile = FSO.OpenTextFile(fileToRead, ForReading, False)
    Set writeFile = FSO.CreateTextFile(fileToWrite, True, False)
    
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
                    Logger.LogDebug "FileHelper.TrimSourceFile", "Readline " & CStr(l) & " . Text: " & ln
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
    Set FSO = Nothing
End Function