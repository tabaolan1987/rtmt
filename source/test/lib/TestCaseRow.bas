Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private mTestCase As ITestCase, mTime As Long

Public Sub Init(TestCase As ITestCase, time As Long)
    Set mTestCase = TestCase
    mTime = time
End Sub

Public Property Get TestCase() As ITestCase
    Set TestCase = mTestCase
End Property

Public Property Get time() As Long
    time = mTime
End Property