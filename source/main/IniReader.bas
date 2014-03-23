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

Public Function init(filePath As String)
  mFilePath = filePath
End Function

Public Function ReadLongKey(sect As String, Keyname As String) As Long
    Dim tmpStr As String
    tmpStr = ReadKey(sect, Keyname)
    If Len(tmpStr) <> 0 Then
        On Error GoTo OnError
        ReadLongKey = CLng(tmpStr)
    Else
        ReadLongKey = 0
    End If
    Exit Function
OnError:
    ReadLongKey = 0
End Function

Public Function ReadIntKey(sect As String, Keyname As String) As Integer
    Dim tmpStr As String
    tmpStr = ReadKey(sect, Keyname)
    If Len(tmpStr) <> 0 Then
        On Error GoTo OnError
        ReadIntKey = CInt(tmpStr)
    Else
        ReadIntKey = 0
    End If
    Exit Function
OnError:
    ReadIntKey = 0
End Function

Public Function ReadBooleanKey(sect As String, Keyname As String) As Boolean
    Dim tmpStr As String
    On Error GoTo OnError
    tmpStr = ReadKey(sect, Keyname)
    If StringHelper.IsEqual(tmpStr, "yes", True) Then
        ReadBooleanKey = True
    Else
        ReadBooleanKey = False
    End If
    Exit Function
OnError:
    ReadBooleanKey = False
End Function

Public Function ReadKey(ByVal sect As String, ByVal Keyname As String) As String
    Dim Worked As Long
    Dim RetStr As String * 128
    Dim StrSize As Long
    Dim iNoOfCharInIni As Long
    Dim sIniString As String
  iNoOfCharInIni = 0
  sIniString = ""
  If sect = "" Or Keyname = "" Then
    'Logger.LogError "IniReader.ReadKey", "Section Or Key To Read Not Specified !!!", Nothing
  Else
    RetStr = Space(128)
    StrSize = Len(RetStr)
    Worked = GetPrivateProfileString(sect, Keyname, "", RetStr, StrSize, mFilePath)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Left$(RetStr, Worked)
    End If
  End If
  sIniString = Trim(sIniString)
  'Logger.LogDebug "IniReader.ReadKey", "Section: " & Sect & ". Key: " & Keyname & ". Value: " & sIniString
  ReadKey = sIniString
End Function

Public Function WriteKey(ByVal sect As String, ByVal Keyname As String, ByVal Wstr As String) As String
Dim Worked As Long
    Dim iNoOfCharInIni As Long
    Dim sIniString As String
  iNoOfCharInIni = 0
  sIniString = ""
  If sect = "" Or Keyname = "" Then
    'Logger.LogError "IniReader.WriteKey", "Section Or Key To Write Not Specified !!!", Nothing
    'MsgBox "Section Or Key To Write Not Specified !!!", vbExclamation, "INI"
  Else
    Worked = WritePrivateProfileString(sect, Keyname, Wstr, mFilePath)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Wstr
    End If
    WriteKey = sIniString
  End If
End Function