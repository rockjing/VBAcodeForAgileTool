Sub Test() 
    Dim msg As Variant 
    msg = GetText() 
    ExtractData (msg)
End Sub
Public Function ExtractData(text As String)
    ' extract

End Function 
Public Function GetText() As String
    Dim JiraAuth 
    Set JiraAuth = CreateObject("MSXML2.XMLHTTP.6.0")
    Dim sUsername As Variant
    Dim sPassword As Variant
    Dim sEncbase64Auth As Variant
    Dim msg As Variant
    sUsername = "yourusername"  
    sPassword = "yourpassword"
    sEncbase64Auth = EncodeBase64(sUsername & ":" & sPassword)
    GetText = ""
    With JiraAuth
        .Open "GET", "http://yourjira/rest/api/", False
        .setRequestHeader "Accept-Language", "en-US,en,q=0.8,zh-cn,q=0.7"
        .setRequestHeader "Accept-Encoding", "gzip,deflate"
        .setRequestHeader "Authorization", "Basic " & sEncbase64Auth
        .send
        GetText = .responseText
    End With
End Function
Public Function EncodeBase64(text As String) As String
    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
   
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
  
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text
  
    Set objNode = Nothing
    Set objXML = Nothing
End Function