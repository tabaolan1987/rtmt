Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private mTestCase As ITestCase, mTime As Long

Public Sub Init(TestCase As ITestCase, Time As Long)
    Set mTestCase = TestCase
    mTime = Time
End Sub

Public Property Get TestCase() As ITestCase
    Set TestCase = mTestCase
End Property

Public Property Get Time() As Long
    Time = mTime
End Property