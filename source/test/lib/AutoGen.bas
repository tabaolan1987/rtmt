Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CountOfLines(Module As CodeModule) As Long
    CountOfLines = Module.CountOfLines
End Function

Public Function Lines(Module As CodeModule, StartLine As Long, count As Long) As String
    Lines = Module.Lines(StartLine, count)
End Function

Public Function SuiteMethodBody(TestMethodNames As Collection) As String
    Dim MethodName As Variant
    SuiteMethodBody = SuiteHeader()
    For Each MethodName In TestMethodNames
        SuiteMethodBody = SuiteMethodBody & vbCrLf & SuiteLine(CStr(MethodName))
    Next
End Function

Public Function SuiteMethod() As String
    SuiteMethod = "Private Function ITest_Suite() As TestSuite"
End Function

Public Function SuiteHeader() As String
    SuiteHeader = "    Set ITest_Suite = New TestSuite"
End Function

Public Function SuiteLine(MethodName As String) As String
    SuiteLine = "    ITest_Suite.AddTest ITest_Manager.ClassName, " & QW(MethodName)
End Function

Public Function EndFunction() As String
    EndFunction = "End Function"
End Function

Public Function GetMethodBody(className As String, MethodName As String) As String
    Dim Module As CodeModule, StartLine As Long, LineCount As Long, BodyLine As Long
    Set Module = GetCodeModule(className)
    If Module Is Nothing Then Exit Function
    'On Error GoTo NO_METHOD
    StartLine = BodyStartLine(Module, MethodName)
    LineCount = BodyLineCount(Module, MethodName)
    If StartLine = 0 Or LineCount = 0 Then Exit Function
    GetMethodBody = Module.Lines(StartLine, LineCount)
'NO_METHOD:
'    Exit Function
End Function

Private Function BodyStartLine(Module As CodeModule, MethodName As String) As Long
    Dim BodyLine As Long, NumFuncLines As Long
    NumFuncLines = 1
    If Module Is Nothing Then Exit Function
    On Error GoTo NO_METHOD
    BodyLine = Module.ProcBodyLine(MethodName, vbext_pk_Proc)
    BodyStartLine = BodyLine + NumFuncLines
NO_METHOD:
    Exit Function
End Function

Public Function BodyLineCount(Module As CodeModule, MethodName As String) As Long
    Dim StartLine As Long, LineCount As Long, BodyLine As Long, NumFuncLines As Long, NumEndFuncLines As Long, BStartLine As Long
    NumFuncLines = 1
    NumEndFuncLines = 1
    If Module Is Nothing Then Exit Function
    On Error GoTo NO_METHOD
    StartLine = Module.ProcStartLine(MethodName, vbext_pk_Proc)
    LineCount = Module.ProcCountLines(MethodName, vbext_pk_Proc)
    Do While Module.Lines(StartLine + LineCount - NumEndFuncLines, 1) = ""
        NumEndFuncLines = NumEndFuncLines + 1
    Loop
    BodyLine = Module.ProcBodyLine(MethodName, vbext_pk_Proc)
    BodyLineCount = LineCount - (BodyLine - StartLine) - NumFuncLines - NumEndFuncLines
    BStartLine = BodyStartLine(Module, MethodName)
NO_METHOD:
    Exit Function
End Function


Public Function GetCodeModule(className) As CodeModule
    Dim Components As VBComponents, Component As VBComponent
    Set Components = Application.VBE.ActiveVBProject.VBComponents
    For Each Component In Components
        If Component.Name = className Then
            Set GetCodeModule = Component.CodeModule
            Exit Function
        End If
    Next
End Function

Public Sub ReplaceMethodBody(className As String, MethodName As String, NewMethodBody As String)
    Dim Module As CodeModule, LineStart As Long, LineCount As Long, BodyLine As Long
    Set Module = GetCodeModule(className)
    If Module Is Nothing Then Exit Sub
    'On Error GoTo NO_METHOD
    DeleteMethodBody Module, MethodName
    InsertMethodBody Module, MethodName, NewMethodBody
NO_METHOD:
    Exit Sub
End Sub

Public Sub DeleteMethodBody(Module As CodeModule, MethodName As String)
    Dim StartLine As Long, LineCount As Long
    If Module Is Nothing Then Exit Sub
    'On Error GoTo NO_METHOD
    StartLine = BodyStartLine(Module, MethodName)
    LineCount = BodyLineCount(Module, MethodName)
    If StartLine = 0 Or LineCount = 0 Then Exit Sub
    Module.DeleteLines StartLine, LineCount
'NO_METHOD:
'    Exit Sub
End Sub

Private Sub InsertMethodBody(Module As CodeModule, MethodName As String, MethodBody As String)
    Dim StartLine As Long
    If Module Is Nothing Then Exit Sub
    'On Error GoTo NO_METHOD
    StartLine = BodyStartLine(Module, MethodName)
    If StartLine = 0 Then Exit Sub
    Module.InsertLines StartLine, MethodBody
'NO_METHOD:
'    Exit Sub
End Sub

Public Function GetTestMethods(className As String) As Collection
    Dim Module As CodeModule, LineNum As Long
    Set GetTestMethods = New Collection
    Set Module = GetCodeModule(className)
    If Module Is Nothing Then Exit Function
    For LineNum = 1 To Module.CountOfLines
        If IsTestMethodLine(Module.Lines(LineNum, 1)) Then
            GetTestMethods.Add Module.ProcOfLine(LineNum, vbext_pk_Set)
        End If
    Next
End Function

Private Function IsTestMethodLine(line As String) As Boolean
    IsTestMethodLine = Left(line, 15) Like "Public Sub Test"
End Function

Public Sub MakeSuite(className As String)
    ReplaceMethodBody className, "ITest_Suite", SuiteMethodBody(GetTestMethods(className))
End Sub

Public Function RunTestLine(MethodName As String) As String
    RunTestLine = "        Case " & QW(MethodName) & ": " & MethodName
End Function

Public Function RunTestHeader() As String
    RunTestHeader = "    Select Case mManager.MethodName"
End Function

Public Function RunTestFooter() As String
    RunTestFooter = "        Case Else: mAssert.Should False, " & QW("Invalid test name: ") & " & mManager.MethodName" & vbCrLf & _
                    "    End Select"
End Function

Public Function RunTestMethodBody(TestMethodNames As Collection) As String
    Dim MethodName As Variant
    RunTestMethodBody = RunTestHeader()
    For Each MethodName In TestMethodNames
        RunTestMethodBody = RunTestMethodBody & vbCrLf & RunTestLine(CStr(MethodName))
    Next
    RunTestMethodBody = RunTestMethodBody & vbCrLf & RunTestFooter()
End Function

Public Sub MakeRunTest(className As String)
    ReplaceMethodBody className, "ITestCase_RunTest", RunTestMethodBody(GetTestMethods(className))
End Sub

Public Sub Prep(Optional className As String)
    Dim Classes As Collection, Name As Variant
    MakeTestClassLister
    If className = "" Then
        Set Classes = TestClasses()
    Else
        Set Classes = New Collection
        Classes.Add className
    End If
    For Each Name In Classes
        MakeSuite CStr(Name)
        MakeRunTest CStr(Name)
    Next
End Sub

Public Function TestClasses() As Collection
    Dim Components As VBComponents, Component As VBComponent
    Set TestClasses = New Collection
    Set Components = Application.VBE.ActiveVBProject.VBComponents
    For Each Component In Components
        If IsClassModule(Component.Type) And IsTestClassName(Component.Name) Then
            TestClasses.Add Component.Name
        End If
    Next
End Function

Public Function IsTestClassName(ComponentName As String) As Boolean
    If Len(ComponentName) <= 6 Then Exit Function
    IsTestClassName = Right(ComponentName, 6) Like "Tester"
End Function

Public Function IsClassModule(ComponentType As vbext_ComponentType) As Boolean
    IsClassModule = (ComponentType = vbext_ct_ClassModule)
End Function

Public Function TestClassHeader() As String
    TestClassHeader = "    Select Case TestClassName"
End Function

Public Function TestClassLine(TestClassName As String) As String
    TestClassLine = "        Case " & QW(TestClassName) & ": Set SelectTestClass = New " & TestClassName
End Function

Public Function TestClassFooter() As String
    TestClassFooter = "        Case Else:" & vbCrLf & _
                      "    End Select"
End Function

Public Function TestClassMethodBody(TestClassNames As Collection) As String
    Dim TestClassName As Variant
    TestClassMethodBody = TestClassHeader()
    For Each TestClassName In TestClassNames
        TestClassMethodBody = TestClassMethodBody & vbCrLf & TestClassLine(CStr(TestClassName))
    Next
    TestClassMethodBody = TestClassMethodBody & vbCrLf & TestClassFooter()
End Function

Public Sub MakeTestClassLister()
    ReplaceMethodBody "TestClassLister", "SelectTestClass", TestClassMethodBody(TestClasses())
End Sub