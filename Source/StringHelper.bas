Option Compare Database

Public Function encodeXml(entry As String) As String
    Dim returnVal As String
    returnVal = entry
    returnVal = Replace(returnVal, "&", "&amp;")
    returnVal = Replace(returnVal, """", "&quot;")
    returnVal = Replace(returnVal, "'", "&apos;")
    returnVal = Replace(returnVal, "<", "&lt;")
    returnVal = Replace(returnVal, ">", "&gt;")
    encodeXml = returnVal
End Function