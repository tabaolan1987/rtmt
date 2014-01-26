Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name, Run, Count
Public Property Let className(RHS As String)

End Property
Public Property Get className() As String

End Property
Public Function Run(test As ITest, Optional RunManager As IRunManager) As IRunManager

End Function
Public Function CountTestCases(test As ITest) As Long

End Function