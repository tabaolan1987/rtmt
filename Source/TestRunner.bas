Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Run(Optional TestClassName As String)
    Dim S As TestSuite, T As ITest, RM As IRunManager, TL As TestClassLister
    Dim C As Collection, VName As Variant
    Set TL = New TestClassLister
    If TestClassName <> "" Then
        Set C = New Collection
        C.Add TestClassName
    Else
        Set C = TL.TestClasses()
    End If
    Set S = New TestSuite
    For Each VName In C
        S.AddTest CStr(VName)
    Next
    Set T = S
    Set RM = New RunManager
    T.Manager.Run T, RM
    RM.Report
End Sub