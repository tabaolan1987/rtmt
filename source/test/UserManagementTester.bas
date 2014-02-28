Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

Private Sub Class_Initialize()
    Set mManager = New TestCaseManager
End Sub

Private Property Get ITestCase_Manager() As TestCaseManager
    Set ITestCase_Manager = mManager
End Property

Private Property Get ITest_Manager() As ITestManager
    Set ITest_Manager = mManager
End Property

Private Sub ITestCase_SetUp(Assert As IAssert)
    Set mAssert = Assert
End Sub

Private Sub ITestCase_TearDown()

End Sub

Public Sub TestCurrentUser()
    On Error GoTo OnError
    Dim cUser As New CurrentUser
    cUser.Init "Carld0"
'mAssert.Equals cUser.Auth, True
   ' mAssert.Equals cUser.Valid, True
OnExit:
    ' finally
    Exit Sub
OnError:
'    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "UserManagementTeser.TestCurrentUser", "", Err
    Resume OnExit
End Sub


Public Sub TestCheckConflict()
    On Error GoTo OnError
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    Dim um As New UserManagement
    um.Init ss
    um.CheckConflict
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "UserManagementTeser.TestCheckConflict", "", Err
    Resume OnExit
End Sub

Public Sub TestCheckDuplicate()
    On Error GoTo OnError
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    Dim um As New UserManagement
    um.Init ss
    um.CheckDuplicate
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "UserManagementTeser.TestCheckDuplicate", "", Err
    Resume OnExit
End Sub

Public Sub TestResolveLdapNotFounds()
    On Error GoTo OnError
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    Dim um As New UserManagement
    um.Init ss
    um.ResolveLdapNotFound
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "UserManagementTeser.TestResolveLdapNotFound", "", Err
    Resume OnExit
End Sub

Public Sub TestResolveLdapConflict()
    On Error GoTo OnError
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    Dim um As New UserManagement
    um.Init ss
    um.ResolveLdapConflict
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "UserManagementTeser.TestResolveLdapConflict", "", Err
    Resume OnExit
End Sub

Public Sub TestResolveUserDataDuplicate()
    On Error GoTo OnError
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    Dim um As New UserManagement
    um.Init ss
    um.ResolveUserDataDuplicate
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "UserManagementTeser.TestResolveUserDataDuplicate", "", Err
    Resume OnExit
End Sub

Public Sub TestResolveUserDataConflict()
    On Error GoTo OnError
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    Dim um As New UserManagement
    um.Init ss
    um.ResolveUserDataConflict
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "UserManagementTeser.TestResolveUserDataConflict", "", Err
    Resume OnExit
End Sub


Public Sub TestMergeUserData()
    On Error GoTo OnError
    Dim ss As SystemSetting
    Set ss = Session.Settings()
    Dim um As New UserManagement
    um.Init ss
    um.MergeUserData
OnExit:
    ' finally
    Exit Sub
OnError:
    mAssert.Should False, Logger.GetErrorMessage("", Err)
    Logger.LogError "UserManagementTeser.TestMergeUserData", "", Err
    Resume OnExit
End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.className, "TestResolveLdapNotFounds"
    ITest_Suite.AddTest ITest_Manager.className, "TestResolveLdapConflict"
    ITest_Suite.AddTest ITest_Manager.className, "TestCheckDuplicate"
    ITest_Suite.AddTest ITest_Manager.className, "TestResolveUserDataDuplicate"
    ITest_Suite.AddTest ITest_Manager.className, "TestCheckConflict"
    ITest_Suite.AddTest ITest_Manager.className, "TestResolveUserDataConflict"
    ITest_Suite.AddTest ITest_Manager.className, "TestMergeUserData"
    ITest_Suite.AddTest ITest_Manager.className, "TestCurrentUser"
    
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.MethodName
        Case "TestResolveLdapNotFounds": TestResolveLdapNotFounds
        Case "TestResolveLdapConflict": TestResolveLdapConflict
        Case "TestCheckDuplicate": TestCheckDuplicate
        Case "TestResolveUserDataDuplicate": TestResolveUserDataDuplicate
        Case "TestCheckConflict": TestCheckConflict
        Case "TestResolveUserDataConflict": TestResolveUserDataConflict
        Case "TestMergeUserData": TestMergeUserData
        Case "TestCurrentUser": TestCurrentUser
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.MethodName
    End Select
End Sub