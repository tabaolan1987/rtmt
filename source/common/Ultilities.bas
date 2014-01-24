'@author Hai Lu
' General utilities function
Option Explicit
#If VBA7 Then
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

Public Function IfTableExists(tblName As String) As Boolean
    'ADO Method
    Dim obj As AccessObject
    Dim dbs As Object
    Set dbs = Application.CurrentData
    IfTableExists = False
    For Each obj In dbs.AllTables
        If obj.Name = tblName Then
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