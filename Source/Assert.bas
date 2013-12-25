Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IAssert
Implements IResultUser

Private mResult As TestResult

Private Sub AddFailure(TestCase As ITestCase, Comment As String)
    mResult.AddFailure TestCase, Comment
    'Debug.Print "Test failed: " & Comment
End Sub

Private Sub IAssert_Delta(Actual As Variant, Expected As Variant, Delta As Variant, Optional Comment As String)

End Sub

Private Sub IAssert_Equals(Actual As Variant, Expected As Variant, Optional Comment As String)
    If Actual <> Expected Then AddFailure CurrentTestCase, NotEqualsComment(Comment, Actual, Expected)
End Sub

Private Sub IAssert_Should(Condition As Boolean, Optional Comment As String)
    If Not Condition Then AddFailure CurrentTestCase, Comment
End Sub

Private Property Get CurrentTestCase() As ITestCase
    Set CurrentTestCase = mResult.CurrentTestCase
End Property

Private Property Set IResultUser_Result(RHS As TestResult)
    Set mResult = RHS
End Property

Private Property Get IResultUser_Result() As TestResult
    Set IResultUser_Result = mResult
End Property

Private Function NotEqualsComment(Comment As String, Actual As Variant, Expected As Variant) As String
    NotEqualsComment = Comment & ":" & " expected: " & QW(CStr(Expected)) & " but was: " & QW(CStr(Actual))
End Function