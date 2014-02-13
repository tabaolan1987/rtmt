Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private mNtid As String
Private mFullName As String
Private mFuncRegion As FunctionRegion
Private mListFuncRg As Collection
Private mAuth As Boolean
Private mValid As Boolean

Public Function Init(iNtid As String, _
                        Optional ss As SystemSetting)
    Dim mData As String
    Dim fields As String
    mNtid = iNtid
    Dim result As String
    Dim tmpNtid As String
    Dim check As String
    Dim tmpDict As Scripting.Dictionary, checkList As Collection
    If ss Is Nothing Then
        Set ss = Session.Settings()
    End If
    
    If ss.EnableTesting Then
        mAuth = True
    Else
        mAuth = False
        For i = 0 To ss.validatorMapping.Count - 1
            fields = fields & ss.validatorMapping.Items(i) & ","
        Next i
        If StringHelper.EndsWith(fields, ",", True) Then
            fields = Left(fields, Len(fields) - 1)
        End If
        mData = "token=" & StringHelper.EncodeURL(ss.Token) _
                & "&fields=" & StringHelper.EncodeURL(fields) _
                & "&ntids=" & StringHelper.EncodeURL(mNtid)
        Logger.LogDebug "CurrentUser.Init", "Post valid ntid: " & mNtid
        result = HttpHelper.PostData(ss.ValidatorURL, mData)
        Logger.LogDebug "CurrentUser.Init", "Result: " & result
        
        If Len(result) > 0 Then
            If StringHelper.IsContain(result, "}", True) And StringHelper.IsContain(result, "{", True) Then
                Set checkList = JSONHelper.parse(result)
                For Each tmpDict In checkList
                    tmpNtid = tmpDict.Item("ntid")
                    check = tmpDict.Item("isvalid")
                    If StringHelper.IsEqual(tmpNtid, mNtid, True) And StringHelper.IsEqual(check, "true", True) Then
                        mAuth = True
                    End If
                Next
            End If
        End If
    End If
    If Ultilities.CheckTables(Constants.SYNC_TYPE_ROLE) Then
        mValid = False
        Dim dbm As New DbManager
        Dim frg As FunctionRegion
        Dim query As String
        Dim data As New Scripting.Dictionary
        Dim RegionName As String
        Dim functionId As String
        Dim roleName As String
        Dim lastRegionName As String
        Dim lastFunctionId As String
        Dim lastRoleName As String
        data.Add Constants.Q_KEY_VALUE, mNtid
        dbm.Init
        query = FileHelper.ReadQuery(Constants.TABLE_USER_PRIVILEGES, Constants.Q_SELECT)
        query = StringHelper.GenerateQuery(query, data)
        Logger.LogDebug "CurrentUser.Init", query
        dbm.OpenRecordSet query
        If Not (dbm.RecordSet.EOF And dbm.RecordSet.BOF) Then
            dbm.RecordSet.MoveFirst
            mValid = True
            Logger.LogDebug "CurrentUser.Init", "Valid: " & mValid
            Set mListFuncRg = New Collection
            Do Until dbm.RecordSet.EOF = True
                RegionName = dbm.GetFieldValue(dbm.RecordSet, "RegionName")
                functionId = dbm.GetFieldValue(dbm.RecordSet, "Function ID")
                roleName = dbm.GetFieldValue(dbm.RecordSet, "roleName")
                Logger.LogDebug "CurrentUser.Init", functionId & " | " & lastFunctionId
                Logger.LogDebug "CurrentUser.Init", roleName & " | " & lastRoleName
                Logger.LogDebug "CurrentUser.Init", RegionName & " | " & lastRegionName
                If (Not StringHelper.IsEqual(RegionName, lastRegionName, True)) _
                    Or (Not StringHelper.IsEqual(functionId, lastFunctionId, True)) Then
                    Logger.LogDebug "CurrentUser.Init", "Init new function region"
                    Set frg = New FunctionRegion
                    frg.Init RegionName, _
                        dbm.GetFieldValue(dbm.RecordSet, "nameFunction"), _
                        roleName, _
                        dbm.GetFieldValue(dbm.RecordSet, "permission"), _
                        functionId
                Else
                    Logger.LogDebug "CurrentUser.Init", "Add more role " & roleName
                    frg.AddRole roleName
                End If
                If (Not StringHelper.IsEqual(RegionName, lastRegionName, True)) _
                    Or (Not StringHelper.IsEqual(functionId, lastFunctionId, True)) Then
                    Logger.LogDebug "CurrentUser.Init", "Add to list "
                    If mFuncRegion Is Nothing Then
                        Set mFuncRegion = frg
                    End If
                    mListFuncRg.Add frg
                End If
                lastRegionName = RegionName
                lastFunctionId = functionId
                lastRoleName = roleName
                dbm.RecordSet.MoveNext
            Loop
            Logger.LogDebug "CurrentUser.Init", "complete..."
        Else
            Logger.LogDebug "CurrentUser.Init", "Not valid"
        End If
    End If
End Function

Public Function ListRegions() As Collection
    Dim frg As FunctionRegion
    Dim list As New Collection
    Dim v As Variant
    Dim check As Boolean
    Dim tmpName As String
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            tmpName = frg.region
            check = False
            For Each v In list
                If StringHelper.IsEqual(CStr(v), tmpName, True) Then
                    check = True
                    Exit For
                End If
            Next v
            If Not check Then
                list.Add tmpName
            End If
        Next frg
    End If
    Set ListRegions = list
End Function

Public Function ListFunctions(region As String) As Collection
    Dim frg As FunctionRegion
    Dim list As New Collection
    Dim v As Variant
    Dim tmpName As String
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            tmpName = frg.region
            If StringHelper.IsEqual(tmpName, region, True) Then
                list.Add frg.Name
            End If
        Next frg
    End If
    Set ListFunctions = list
End Function


Public Function SelectFunc(fName As String)
    Dim frg As FunctionRegion
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            If StringHelper.IsEqual(frg.value, fName, True) Then
                Set mFuncRegion = frg
                Exit For
            End If
        Next frg
    End If
End Function

Public Property Get Valid() As Boolean
    Valid = mValid
End Property

Public Property Get Auth() As Boolean
    Auth = mAuth
End Property

Public Property Get ntid() As String
    ntid = mNtid
End Property

Public Property Get FuncRegion() As FunctionRegion
    Set FuncRegion = mFuncRegion
End Property

Public Property Get ListFuncRg() As Collection
    Set ListFuncRg = mListFuncRg
End Property

Public Function SetNtid(ntid As String)
    mNtid = ntid
End Function