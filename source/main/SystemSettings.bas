Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
' To read & write Settings.ini
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#Else
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Private mServerName As String
Private mDatabaseName As String
Private mPort As String
Private mUsername As String
Private mPassword As String
Private mSyncTables() As String
Private mSyncUsers As Scripting.Dictionary
Private mLineToRemove() As Integer

Private Function IniFileName() As String
  IniFileName = FileHelper.CurrentDbPath & "config\settings.ini"
End Function

Private Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String) As String
    Dim Worked As Long
    Dim RetStr As String * 128
    Dim StrSize As Long
    Dim iNoOfCharInIni As Long
    Dim sIniString As String
  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    Logger.LogError "Utilities.ReadIniFileString", "Section Or Key To Read Not Specified !!!", Nothing
    MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
  Else
    RetStr = Space(128)
    StrSize = Len(RetStr)
    Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Left$(RetStr, Worked)
    End If
  End If
  ReadIniFileString = sIniString
End Function

Private Function WriteIniFileString(ByVal Sect As String, ByVal Keyname As String, ByVal Wstr As String) As String
Dim Worked As Long
    Dim iNoOfCharInIni As Long
    Dim sIniString As String
  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    Logger.LogError "Utilities.WriteIniFileString", "Section Or Key To Write Not Specified !!!", Nothing
    'MsgBox "Section Or Key To Write Not Specified !!!", vbExclamation, "INI"
  Else
    Worked = WritePrivateProfileString(Sect, Keyname, Wstr, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Wstr
    End If
    WriteIniFileString = sIniString
  End If
End Function

Public Function Init()
    mServerName = ReadIniFileString(Constants.SECTION_REMOTE_DATABASE, Constants.KEY_SERVER_NAME)
    mDatabaseName = ReadIniFileString(Constants.SECTION_REMOTE_DATABASE, Constants.KEY_DATABASE_NAME)
    mPort = ReadIniFileString(Constants.SECTION_REMOTE_DATABASE, Constants.KEY_PORT)
    mUsername = ReadIniFileString(Constants.SECTION_REMOTE_DATABASE, Constants.KEY_USERNAME)
    mPassword = ReadIniFileString(Constants.SECTION_REMOTE_DATABASE, Constants.KEY_PASSWORD)
    
    mSyncTables = FileHelper.ReadSSFile(FileHelper.CurrentDbPath & Constants.SS_SYNC_TABLES)
    
    Dim source As String, tmpList() As String, ln As String, arraySize As Integer, i As Integer
    source = ReadIniFileString(Constants.SECTION_USER_DATA, Constants.KEY_LINE_TO_REMOVE)
    tmpList = Split(source, ",")
    For i = LBound(tmpList) To UBound(tmpList)
        ln = Trim(tmpList(i))
        If Len(ln) <> 0 Then
            ReDim Preserve mLineToRemove(arraySize)
            mLineToRemove(arraySize) = CInt(ln)
            arraySize = arraySize + 1
        End If
    Next
    Dim tl() As String
    Set mSyncUsers = New Scripting.Dictionary
    tmpList = FileHelper.ReadSSFile(FileHelper.CurrentDbPath & Constants.SS_SYNC_USERS)
    For i = LBound(tmpList) To UBound(tmpList)
        ln = Trim(tmpList(i))
        If Len(ln) <> 0 Then
            tl = Split(ln, ":")
            mSyncUsers.Add Trim(tl(0)), Trim(tl(1))
        End If
    Next
End Function

Public Property Get ServerName() As String
    ServerName = mServerName
End Property

Public Property Get DatabaseName() As String
    DatabaseName = mDatabaseName
End Property

Public Property Get Port() As String
    Port = mPort
End Property

Public Property Get Username() As String
    Username = mUsername
End Property

Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Get SyncTables() As String()
    SyncTables = mSyncTables
End Property

Public Property Get SyncUsers() As Scripting.Dictionary
    Set SyncUsers = mSyncUsers
End Property

Public Property Get LineToRemove() As Integer()
    LineToRemove = mLineToRemove
End Property