Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mRegion As String
Private mName As String
Private mRole As Collection
Private mFuncRgID As String
Private mPermission As Collection

Public Function Init(iRegion As String, _
                        iRole As String, _
                        iPermission As String)
    Logger.LogDebug "FunctionRegion.Init", "Region: " & iRegion & " Role: " & iRole & " Permission: " & iPermission
    mRegion = iRegion
    AddRole iRole
    AddPermission iPermission
End Function

Public Function AddRole(role As String)
    If mRole Is Nothing Then
        Set mRole = New Collection
    End If
    mRole.Add role
End Function

Public Function AddPermission(permission As String)
    If mPermission Is Nothing Then
        Set mPermission = New Collection
    End If
    mPermission.Add permission
End Function

Public Property Get value() As String
    value = mRegion ' & " - " & mName
End Property

Public Function SetFuncRgId(id As String)
    mFuncRgID = id
End Function

Public Property Get FuncRgID() As String
    FuncRgID = mFuncRgID
End Property

Public Property Get Region() As String
    Region = mRegion
End Property

Public Function SetFuncName(iName As String)
    mName = iName
End Function

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get role() As Collection
    Set role = mRole
End Property

Public Property Get permission() As Collection
    Set permission = mPermission
End Property