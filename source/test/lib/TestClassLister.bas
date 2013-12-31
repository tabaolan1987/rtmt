Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Public Function NewTestClass(TestClassName As String) As ITest
    Dim test As ITest
    Set NewTestClass = SelectTestClass(TestClassName)
    If NewTestClass Is Nothing Then Exit Function
    NewTestClass.Manager.className = TestClassName
End Function

Public Function SelectTestClass(TestClassName As String) As ITest
    Select Case TestClassName
        Case "DbManagerTester": Set SelectTestClass = New DbManagerTester
        Case "FileHelperTester": Set SelectTestClass = New FileHelperTester
        Case "StringHelperTester": Set SelectTestClass = New StringHelperTester
        Case "ReportingTester": Set SelectTestClass = New ReportingTester
        Case Else:
    End Select
End Function