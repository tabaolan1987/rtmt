Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @author Hai Lu
' Report meta data object
Option Explicit
Private mQuery As String
Private mValid As Boolean

Public Function Init(name As String)
    mQuery = FileHelper.readFile(Constants.RP_ROOT_FOLDER & name & ".sql")
    If Len(mQuery) <> 0 Then
        mValid = True
    End If
End Function

Public Property Get query() As String
    query = mQuery
End Property

Public Property Get Valid() As Boolean
    Valid = mValid
End Property