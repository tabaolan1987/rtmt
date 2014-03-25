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

Public Property Get IsRole(role As String) As Boolean
    Dim v As Variant
    IsRole = False
    If Not mFuncRegion.role Is Nothing Then
        For Each v In mFuncRegion.role
            If StringHelper.IsEqual(CStr(v), role, True) Then
                Logger.LogDebug "CurrentUser.IsRole", "Found role: " & CStr(v)
                IsRole = True
                Exit For
            End If
        Next v
    End If
End Property

Public Function init(iNtid As String, _
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
    ' skip temp
    If Ultilities.CheckTables(Constants.SYNC_TYPE_ROLE) Then
        mValid = False
        Dim dbm As New DbManager
        Dim frg As FunctionRegion
        Dim query As String
        Dim data As New Scripting.Dictionary
        Dim regionName As String
        Dim functionId As String
        Dim roleName As String
        Dim lastRegionName As String
        Dim lastFunctionId As String
        Dim lastRoleName As String
        data.Add Constants.Q_KEY_VALUE, mNtid
        dbm.init
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
                regionName = dbm.GetFieldValue(dbm.RecordSet, "RegionName")
                roleName = dbm.GetFieldValue(dbm.RecordSet, "roleName")
                Logger.LogDebug "CurrentUser.Init", roleName & " | " & lastRoleName
                Logger.LogDebug "CurrentUser.Init", regionName & " | " & lastRegionName
                If (Not StringHelper.IsEqual(regionName, lastRegionName, True)) Then
                    Logger.LogDebug "CurrentUser.Init", "Init new function region " & regionName
                    Set frg = New FunctionRegion
                    frg.init regionName, _
                        roleName, _
                        dbm.GetFieldValue(dbm.RecordSet, "permission")
                Else
                    Logger.LogDebug "CurrentUser.Init", "Add more role " & roleName
                    frg.AddRole roleName
                End If
                If (Not StringHelper.IsEqual(regionName, lastRegionName, True)) Then
                    Logger.LogDebug "CurrentUser.Init", "Add to list "
                    If mFuncRegion Is Nothing Then
                        Set mFuncRegion = frg
                    End If
                    mListFuncRg.Add frg
                End If
                lastRegionName = regionName
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
            tmpName = frg.Region
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

Public Function ListFunctions(Region As String) As Collection
    Dim frg As FunctionRegion
    Dim list As New Collection
    Dim v As Variant
    Dim tmpName As String
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            tmpName = frg.Region
            If StringHelper.IsEqual(tmpName, Region, True) Then
                list.Add frg.Name
            End If
        Next frg
    End If
    Set ListFunctions = list
End Function

Public Function SelectRegion(rname As String)
    Dim frg As FunctionRegion
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            If StringHelper.IsEqual(frg.Region, rname, True) Then
                Set mFuncRegion = frg
                Exit For
            End If
        Next frg
    End If
End Function

Public Function SelectFunc(fname As String)
    Dim frg As FunctionRegion
    If mValid And Not mListFuncRg Is Nothing Then
        For Each frg In mListFuncRg
            If StringHelper.IsEqual(frg.value, fname, True) Then
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