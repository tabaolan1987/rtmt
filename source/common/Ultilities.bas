'@author Hai Lu
' General utilities function
Option Explicit

Private checkInternetFlag As String
Const NoError = 0

#If VBA7 Then
Declare PtrSafe Function WNetGetUser Lib "mpr.dll" _
      Alias "WNetGetUserA" (ByVal lpName As String, _
      ByVal lpUserName As String, lpnLength As Long) As Long
Private Declare PtrSafe Function GetClassNameA Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) _
    As Long
Private Declare PtrSafe Function GetWindow Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal wCmd As Long) _
    As Long
Private Declare PtrSafe Function ShowWindowAsync Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nCmdShow As Long) _
    As Boolean
#Else
Declare Function WNetGetUser Lib "mpr.dll" _
      Alias "WNetGetUserA" (ByVal lpName As String, _
      ByVal lpUserName As String, lpnLength As Long) As Long
Private Declare Function GetClassNameA Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) _
    As Long
Private Declare Function GetWindow Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal wCmd As Long) _
    As Long
Private Declare Function ShowWindowAsync Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nCmdShow As Long) _
    As Boolean
#End If
    
Private mIniReader As IniReader

Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

Public Function SystemIniReader() As IniReader
    If mIniReader Is Nothing Then
        Set mIniReader = New IniReader
        mIniReader.Init FileHelper.CurrentDbPath & Constants.SETTINGS_FILE
    End If
    Set SystemIniReader = mIniReader
End Function

Public Sub MakeAccde()
    Dim sourcedb As String, targetdb As String
    sourcedb = FileHelper.CurrentDbPath & "target\rolemapping.accdb"
    targetdb = FileHelper.CurrentDbPath & "target\rolemapping.accde"
    Logger.LogDebug "Ultilities.MakeAccde", "source db:" & sourcedb
    Logger.LogDebug "Ultilities.MakeAccde", "target db:" & targetdb
    
    Dim AccessApplication As New Access.Application
    With AccessApplication
        .Visible = False
        .AutomationSecurity = 1 'MsoAutomationSecurityLow
        .UserControl = True
        .SysCmd 603, sourcedb, targetdb 'this makes the ACCDE file
        .Quit
    End With
End Sub

Public Sub WDeleteAllAppData()
    WDeleteAllTables
    Dim db As New DbManager
    db.ExecuteServerQuery "delete from user_data"
    db.ExecuteServerQuery "delete from user_data_mapping_role"
    db.ExecuteServerQuery "delete from audit_logs"
    db.ExecuteServerQuery "delete from user_change_log"
End Sub

Public Sub WDeleteAllTables()
    Dim i As Integer
    Dim dbm As New DbManager
    dbm.Init
    Dim tmpStr As String
    Dim tables() As String
    tables = Session.Settings.SyncMappingTables
    If Not Ultilities.IsVarArrayEmpty(tables) Then
        For i = LBound(tables) To UBound(tables)
            tmpStr = tables(i)
            dbm.RecycleTableName tmpStr
        Next i
    End If
    tables = Session.Settings.SyncRoleTables
    If Not Ultilities.IsVarArrayEmpty(tables) Then
        For i = LBound(tables) To UBound(tables)
            tmpStr = tables(i)
            dbm.RecycleTableName tmpStr
        Next i
    End If
    tables = Session.Settings.SyncTables
    If Not Ultilities.IsVarArrayEmpty(tables) Then
        For i = LBound(tables) To UBound(tables)
            tmpStr = tables(i)
            dbm.RecycleTableName tmpStr
        Next i
    End If
    tables = Session.Settings.JunkTables
    If Not Ultilities.IsVarArrayEmpty(tables) Then
        
        For i = LBound(tables) To UBound(tables)
            tmpStr = tables(i)
            dbm.DeleteTable tmpStr
        Next i
    End If
    dbm.RecycleTableName "ChangeLog"
    dbm.RecycleTableName "audit_logs"
    dbm.RecycleTableName "user_change_log"
    dbm.RecycleTableName "USysApplicationLog"
    dbm.RecycleTableName "tmp_curriculum"
    dbm.RecycleTableName "w_in_curriculum_not_in_db"
    dbm.RecycleTableName "w_in_db_not_in_curriculum"
    dbm.RecycleTableName "w_invalid_bluesprint_role"
    dbm.RecycleTableName "w_invalid_specialism"
    dbm.RecycleTableName "w_invalid_standard_function"
    dbm.RecycleTableName "w_invalid_standard_team"
    dbm.RecycleTableName "w_invalid_sub_function"
    dbm.Recycle
End Sub

Public Function IfTableExists(tblName As String) As Boolean
    'ADO Method
    Dim obj As AccessObject
    Dim dbs As Object
    Set dbs = Application.CurrentData
    IfTableExists = False
    For Each obj In dbs.AllTables
        If StringHelper.IsEqual(obj.Name, tblName, True) Then
            IfTableExists = True
            Exit For
        End If
    Next obj
End Function

Function IsVarArrayEmpty(anArray As Variant)

    Dim i As Integer
    
    On Error Resume Next
        i = UBound(anArray, 1)
    If Err.Number = 0 Then
        IsVarArrayEmpty = False
    Else
        IsVarArrayEmpty = True
    End If

End Function



Private Function GetClassName( _
    ByVal hwnd As Long) _
    As String

    Dim lpClassName As String
    Dim lLen As Long

    lpClassName = String(255, 32)
    lLen = GetClassNameA(hwnd, lpClassName, 255)
    If lLen > 0 Then
        GetClassName = Left(lpClassName, lLen)
    End If

End Function

Public Sub ShowDbWindow(ByVal bCmdShow As Boolean)

    Dim hWndApp As Long
    
    hWndApp = GetWindow(Application.hWndAccessApp, GW_CHILD)
    Do Until hWndApp = 0
        If GetClassName(hWndApp) = "MDIClient" Then
            Exit Do
        End If
        hWndApp = GetWindow(hWndApp, GW_HWNDNEXT)
    Loop
    
    If hWndApp > 0 Then
        hWndApp = GetWindow(hWndApp, GW_CHILD)
        Do Until hWndApp = 0
            If GetClassName(hWndApp) = "ODb" Then
                Exit Do
            End If
            hWndApp = GetWindow(hWndApp, GW_HWNDNEXT)
        Loop
    End If
    
    If hWndApp > 0 Then
        ShowWindowAsync hWndApp, IIf(bCmdShow, SW_SHOW, SW_HIDE)
    End If

End Sub

Function ShowToolTip(ShowControl As String)
          Dim MyControl As Control
          Dim MyToolTip As Control
          Dim z As Integer
          Const Separator = 80
          On Error Resume Next

          Set MyControl = Screen.ActiveForm(ShowControl)
          Set MyToolTip = Screen.ActiveForm!ToolTip

          If MyToolTip.Visible = False Then
              MyToolTip = MyControl.Tag
              MyToolTip.Left = MyControl.Left + (Separator * 2)
              MyToolTip.Top = MyControl.Top + MyControl.Height + Separator
              MyToolTip.Visible = True
            
              ' Optional: Display ToolTip on the Status Bar.
              z = SysCmd(SYSCMD_SETSTATUS, MyToolTip.value)
          End If

End Function

Public Function GetUserName() As String
    ' Buffer size for the return string.
    Const lpnLength As Integer = 255
    ' Get return buffer space.
    Dim Status As Integer
    ' For getting user information.
    Dim lpName, lpUserName As String
    ' Assign the buffer size constant to lpUserName.
    lpUserName = Space$(lpnLength + 1)
    ' Get the log-on name of the person using product.
    Status = WNetGetUser(lpName, lpUserName, lpnLength)
    ' See whether error occurred.
    If Status = NoError Then
         ' This line removes the null character. Strings in C are null-
         ' terminated. Strings in Visual Basic are not null-terminated.
         ' The null character must be removed from the C strings to be used
         ' cleanly in Visual Basic.
         lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    Else
         ' An error occurred.
         Logger.LogError "Ultilities.GetUserName", "Unable to get the name.", Nothing
         End
    End If
    ' Display the name of the person logged on to the machine.
    Logger.LogDebug "Ultilities.GetUserName", "The person logged on this machine is: " & lpUserName
    GetUserName = lpUserName

End Function

Public Function CheckTables(mType As Integer) As Boolean
    Dim check As Boolean, _
        SyncTables() As String, _
        prop As SystemSetting, _
        isEmpty As Boolean, _
        stTable As String
    Set prop = Session.Settings()
    Select Case mType
        Case Constants.SYNC_TYPE_DEFAULT:
            SyncTables = prop.SyncTables
        Case Constants.SYNC_TYPE_ROLE:
            SyncTables = prop.SyncRoleTables
        Case Constants.SYNC_TYPE_MAPPING:
            SyncTables = prop.SyncMappingTables
    End Select
    
    isEmpty = Ultilities.IsVarArrayEmpty(SyncTables)
    check = True
    If isEmpty = False Then
        Dim i As Integer
        For i = LBound(SyncTables) To UBound(SyncTables)
            stTable = Trim(SyncTables(i))
            If Not Ultilities.IfTableExists(stTable) Then
                check = False
                Exit For
            End If
        Next i
    End If
    CheckTables = check
End Function

Function IsLoaded(ByVal strFormName As String) As Boolean
' Returns True if the specified form is open in Form view or Datasheet view.
' Use form name according to Access, not VBA.
' Only works for Access
    Dim oAccessObject As AccessObject

    Set oAccessObject = CurrentProject.AllForms(strFormName)
    If oAccessObject.IsLoaded Then
        If oAccessObject.CurrentView <> acCurViewDesign Then
            IsLoaded = True
        End If
    End If

End Function

Function CheckInternetConnection() As Boolean
    CheckInternetConnection = True
    
End Function


Function IsReadOnly() As Boolean
    If CurrentDb.Updatable Then
        IsReadOnly = False
    Else
        IsReadOnly = True
    End If
End Function

Public Sub WGetAllTables()
    GetTables Session.Settings.SyncMappingTables
    GetTables Session.Settings.SyncRoleTables
    GetTables Session.Settings.SyncTables
End Sub

Function GetTables(SyncTables() As String)
    On Error GoTo OnError
    Dim dbm As New DbManager, _
    prop As SystemSetting, _
            isEmpty As Boolean, _
            stTable As String
    Dim sh As SyncHelper
    
    isEmpty = Ultilities.IsVarArrayEmpty(SyncTables)
    If Not isEmpty Then
        Dim i As Integer
        dbm.RecycleTableName Constants.TABLE_SYNC_CONFLICT
        For i = LBound(SyncTables) To UBound(SyncTables)
            stTable = Trim(SyncTables(i))
            Set sh = New SyncHelper
            sh.Init stTable
            sh.sync
            sh.Recycle
        Next i
    End If
OnExit:
    Exit Function
OnError:
    Logger.LogError "Ultilities.SyncTables", "An error occurred while processing", Err
    Resume OnExit
End Function


Sub WDisableShift()
    'This function disable the shift at startup. This action causes
    'the Autoexec macro and Startup properties to always be executed.
    
    On Error GoTo errDisableShift
    
    Dim db As DAO.Database
    Dim prop As DAO.Property
    Const conPropNotFound = 3270
    
    Set db = CurrentDb()
    
    'This next line disables the shift key on startup.
    db.Properties("AllowByPassKey") = False
    
    'The function is successful.
    Exit Sub
    
errDisableShift:
    'The first part of this error routine creates the "AllowByPassKey
    'property if it does not exist.
    If Err = conPropNotFound Then
    Set prop = db.CreateProperty("AllowByPassKey", _
    dbBoolean, False)
    db.Properties.Append prop
    Resume Next
    Else
    MsgBox "Function 'ap_DisableShift' did not complete successfully."
    Exit Sub
    End If

End Sub

Sub WEnableShift()
    'This function enables the SHIFT key at startup. This action causes
    'the Autoexec macro and the Startup properties to be bypassed
    'if the user holds down the SHIFT key when the user opens the database.
    
    On Error GoTo errEnableShift
    
    Dim db As DAO.Database
    Dim prop As DAO.Property
    Const conPropNotFound = 3270
    
    Set db = CurrentDb()
    
    'This next line of code disables the SHIFT key on startup.
    db.Properties("AllowByPassKey") = True
    
    'function successful
    Exit Sub
    
errEnableShift:
    'The first part of this error routine creates the "AllowByPassKey
    'property if it does not exist.
    If Err = conPropNotFound Then
    Set prop = db.CreateProperty("AllowByPassKey", _
    dbBoolean, True)
    db.Properties.Append prop
    Resume Next
    Else
    MsgBox "Function 'ap_DisableShift' did not complete successfully."
    Exit Sub
    End If
End Sub