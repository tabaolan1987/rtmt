Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private mRegion As String
Private mName As String
Private mRole As Collection
Private mFuncRgID As String
Private mPermission As String

Public Function Init(iRegion As String, _
                        iName As String, _
                        iRole As String, _
                        iPermission As String, _
                        iFuncRgID As String)
    mRegion = iRegion
    mName = iName
    AddRole iRole
    mPermission = iPermission
    mFuncRgID = iFuncRgID
End Function

Public Function AddRole(role As String)
    If mRole Is Nothing Then
        Set mRole = New Collection
    End If
    mRole.Add role
End Function

Public Property Get value() As String
    value = mRegion & " - " & mName
End Property

Public Property Get FuncRgID() As String
    FuncRgID = mFuncRgID
End Property

Public Property Get region() As String
    region = mRegion
End Property


Public Property Get name() As String
    name = mName
End Property


Public Property Get role() As Collection
    Set role = mRole
End Property

Public Property Get Permission() As String
    Permission = mPermission
End Property