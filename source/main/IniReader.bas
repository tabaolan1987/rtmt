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

Private mFilePath As String

Public Function Init(filePath As String)
  mFilePath = filePath
End Function

Public Function ReadKey(ByVal Sect As String, ByVal Keyname As String) As String
    Dim Worked As Long
    Dim RetStr As String * 128
    Dim StrSize As Long
    Dim iNoOfCharInIni As Long
    Dim sIniString As String
  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    'Logger.LogError "IniReader.ReadKey", "Section Or Key To Read Not Specified !!!", Nothing
  Else
    RetStr = Space(128)
    StrSize = Len(RetStr)
    Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, mFilePath)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Left$(RetStr, Worked)
    End If
  End If
  sIniString = Trim(sIniString)
  'Logger.LogDebug "IniReader.ReadKey", "Section: " & Sect & ". Key: " & Keyname & ". Value: " & sIniString
  ReadKey = sIniString
End Function

Public Function WriteKey(ByVal Sect As String, ByVal Keyname As String, ByVal Wstr As String) As String
Dim Worked As Long
    Dim iNoOfCharInIni As Long
    Dim sIniString As String
  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    'Logger.LogError "IniReader.WriteKey", "Section Or Key To Write Not Specified !!!", Nothing
    'MsgBox "Section Or Key To Write Not Specified !!!", vbExclamation, "INI"
  Else
    Worked = WritePrivateProfileString(Sect, Keyname, Wstr, mFilePath)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Wstr
    End If
    WriteKey = sIniString
  End If
End Function