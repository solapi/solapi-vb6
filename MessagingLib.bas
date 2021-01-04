Attribute VB_Name = "MessagingLib"
Option Explicit

Public Function SHA256(ByVal sTextToHash As String, ByVal sSecretKey As String) As String

    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SecretKey() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA256")
    
    TextToHash = asc.Getbytes_4(sTextToHash)
    SecretKey = asc.Getbytes_4(sSecretKey)
    enc.key = SecretKey
    
    Dim bytes() As Byte
    Dim sig As String
    
    bytes = enc.ComputeHash_2((TextToHash))
    SHA256 = ConvToHexString(bytes)
    
    Set enc = Nothing
    Set asc = Nothing
End Function


Private Function ConvToHexString(vIn As Variant) As Variant
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function
    
Function GetSignature(ApiKey As String, data As String, ApiSecret As String)
    Dim hexStr
    hexStr = SHA256(data, ApiSecret)
    GetSignature = hexStr
End Function

'Salt 생성
Function GetSalt( _
    Optional ByVal Length As Long = 32, _
    Optional charset As String = "abcdefghijklmnopqrstuvwxyz0123456789" _
    ) As String
    Dim chars() As Byte, value() As Byte, chrUprBnd As Long, i As Long
    If Length > 0& Then
        Randomize
        chars = charset
        chrUprBnd = Len(charset) - 1&
        Length = (Length * 2&) - 1&
        ReDim value(Length) As Byte
        For i = 0& To Length Step 2&
            value(i) = chars(CLng(chrUprBnd * Rnd) * 2&)
        Next
    End If
        
    GetSalt = value
End Function

Function GetAuth()
    Dim salt
    Dim dateStr
    Dim data As String
    
    salt = GetSalt()
    dateStr = GetUDTDateTime
    data = dateStr & salt
    
    GetAuth = "HMAC-SHA256 apiKey=" & ApiKey & ", date=" & dateStr & ", salt=" & salt & ", signature=" & GetSignature(ApiKey, data, ApiSecret)
End Function

Public Function Request(path As String, Optional method As String = "GET", Optional data As Dictionary = vbNull) As WebResponse
    Dim Client As New WebClient
    Client.BaseUrl = Protocol & "://" & Domain & Prefix
    
    Dim AuthString
    AuthString = GetAuth()

    Dim req As New WebRequest
    req.Resource = path
    Select Case method
        Case "GET"
            req.method = WebMethod.HttpGet
        Case "POST"
            req.method = WebMethod.HttpPost
    End Select
    
    req.Format = WebFormat.JSON
    req.AddHeader "Authorization", AuthString

    If Not IsNull(data) Then
        Set req.Body = data
    End If

    Dim res As WebResponse
    Set res = Client.Execute(req)
    Dim line As Integer
    Dim indent As Integer
    Debug.Print (res.Content)
    
    Set Request = res
End Function


Function ReadFile(path As String) As Byte()
    On Error GoTo Handler
    Dim fileNum As Integer
    Dim bytes() As Byte

    fileNum = FreeFile
    Debug.Print (path)
    Open path For Binary As fileNum
    ReDim bytes(LOF(fileNum) - 1)
    Get fileNum, , bytes
    Close fileNum
  
    ReadFile = bytes
    Exit Function
Handler:
    Debug.Print "파일을 읽어올 수 없습니다."
    Debug.Print "Error " & Err.Number & Err.Description
    Err.Raise Err.Number, "Source: " & Err.Source, Err.Description
End Function

Function ConvertBytesToBase64(web_Bytes() As Byte)
    ' Use XML to convert to Base64
    Dim web_XmlObj As Object
    Dim web_Node As Object

    Set web_XmlObj = CreateObject("MSXML2.DOMDocument")
    Set web_Node = web_XmlObj.createElement("b64")

    web_Node.DataType = "bin.base64"
    web_Node.nodeTypedValue = web_Bytes
    ConvertBytesToBase64 = RemoveLines(web_Node.Text)

    Set web_Node = Nothing
    Set web_XmlObj = Nothing
End Function

Function RemoveLines(myString As String)
    myString = Replace(myString, vbTab, vbNullString)   ' tab 문자열을 제거
    myString = Replace(myString, Chr(13), vbNullString)
    myString = Replace(myString, Chr(10), vbNullString)
    myString = Replace(myString, vbCrLf, vbNullString)
    myString = Replace(myString, vbNewLine, vbNullString)
    RemoveLines = myString
End Function

Function UploadImage(path As String) As WebResponse
    Dim bytes() As Byte
    Dim b64Image As String

    bytes = ReadFile(path)
    
    Dim i As Integer
    
    b64Image = ConvertBytesToBase64(bytes)
    Debug.Print (Left(b64Image, 100))
    
    Dim Body As New Dictionary
    Body.Add "type", "MMS"
    Body.Add "file", b64Image

    Dim Response As WebResponse
    Set UploadImage = Request("storage/v1/files", "POST", Body)
End Function

Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

Function UploadKakaoImage(path As String, url As String) As WebResponse
    Dim bytes() As Byte
    Dim b64Image As String

    bytes = ReadFile(path)

    b64Image = ConvertBytesToBase64(bytes)
    Dim Body As New Dictionary
    Body.Add "type", "KAKAO"
    Body.Add "file", b64Image
    Body.Add "name", GetFileNameFromPath(path)
    Body.Add "link", url

    Dim Response As WebResponse
    Set UploadKakaoImage = Request("storage/v1/files", "POST", Body)
End Function


