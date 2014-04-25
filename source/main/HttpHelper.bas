Option Explicit
Public Function PostData(url As String, Optional sData As String) As String
    Dim SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    Dim xmlhttp
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.SetOption 2, xmlhttp.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    Dim txtRes As String
    txtRes = ""
    Logger.LogDebug "HttpHelper.PostData", "Url:" & url & ". Data: " & sData
    With xmlhttp
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send (sData)
        txtRes = .responsetext
    End With
    Logger.LogDebug "HttpHelper.PostData", "Result: " & txtRes
    PostData = txtRes
End Function