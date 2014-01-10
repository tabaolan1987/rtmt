Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Run(Optional TestClassName As String)
    Dim s As TestSuite, T As ITest, RM As IRunManager, tl As TestClassLister
    Dim c As Collection, VName As Variant
    Set tl = New TestClassLister
    If TestClassName <> "" Then
        Set c = New Collection
        c.Add TestClassName
    Else
        Set c = tl.TestClasses()
    End If
    Set s = New TestSuite
    For Each VName In c
        s.AddTest CStr(VName)
    Next
    Set T = s
    Set RM = New RunManager
    T.Manager.Run T, RM
    RM.Report
End Sub