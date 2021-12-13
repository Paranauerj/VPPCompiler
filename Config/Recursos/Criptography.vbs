'        @PARANAUERJ DEVELOPEMENT
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







function md5hashBytes(aBytes)
    Dim MD5
    set MD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    MD5.Initialize()
    md5hashBytes = MD5.ComputeHash_2( (aBytes) )
end function

function sha1hashBytes(aBytes)
    Dim sha1
    set sha1 = CreateObject("System.Security.Cryptography.SHA1Managed")

    sha1.Initialize()
    sha1hashBytes = sha1.ComputeHash_2( (aBytes) )
end function

function sha256hashBytes(aBytes)
    Dim sha256
    set sha256 = CreateObject("System.Security.Cryptography.SHA256Managed")

    sha256.Initialize()
    sha256hashBytes = sha256.ComputeHash_2( (aBytes) )
end function

function sha256HMACBytes(aBytes, aKey)
    Dim sha256
    set sha256 = CreateObject("System.Security.Cryptography.HMACSHA256")

    sha256.Initialize()
    sha256.key=aKey
    sha256HMACBytes = sha256.ComputeHash_2( (aBytes) )
end function

function stringToUTFBytes(aString)
    Dim UTF8
    Set UTF8 = CreateObject("System.Text.UTF8Encoding")
    stringToUTFBytes = UTF8.GetBytes_4(aString)
end function

function bytesToHex(aBytes)
    dim hexStr, x
    for x=1 to lenb(aBytes)
        hexStr= hex(ascb(midb( (aBytes),x,1)))
        if len(hexStr)=1 then hexStr="0" & hexStr
        bytesToHex=bytesToHex & hexStr
    next
end function

Function BytesToBase64(varBytes)
    With CreateObject("MSXML2.DomDocument").CreateElement("b64")
        .dataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = .Text
    End With
End Function

Function BytesToBase64UrlEncode(varBytes)
    With CreateObject("MSXML2.DomDocument").CreateElement("b64")
        .dataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64UrlEncode = replace(replace(replace(replace(replace(.Text,chr(13),""),chr(10),""),"+", "-"),"/", "_"),"=", "")
    End With
End Function

Function GetBytes(sPath)
    With CreateObject("Adodb.Stream")
        .Type = 1 
        .Open
        .LoadFromFile sPath
        .Position = 0
        GetBytes = .Read
        .Close
    End With
	End function
	
Function md5(str)
	decoded = bytesToHex(md5hashBytes(stringToUTFBytes(str)))
	md5 = decoded
	End function
	
Function sha256(str)
	decoded = bytesToHex(sha256hashBytes(stringToUTFBytes("Hello World")))
	sha256 = decoded
	End function