Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private mNtid As String
Private mFullName As String
Private mFuncRegion As FunctionRegion
Private mListFuncRg As Collection

Public Property Set NTID(n As String)
    mNtid = n
End Property

Public Property Get NTID() As String
    NTID = mNtid
End Property

Public Property Get FullName() As String
    FullName = mFullName
End Property

Public Property Get FuncRegion() As FunctionRegion
    Set FuncRegion = mFuncRegion
End Property

Public Property Get ListFuncRg() As Collection
    Set ListFuncRg = mListFuncRg
End Property