Option Compare Database

Public Function PostData(url As String, Optional sData As String) As String
    Dim xmlhttp
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.SetOption 2, xmlhttp.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    Dim txtRes As String
    With xmlhttp
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send (sData)
        txtRes = .responsetext
    End With
    PostData = txtRes
End Function