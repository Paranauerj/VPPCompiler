'        @PARANAUERJ DEVELOPEMENT
'
'
' EXTENSÃO PARA WEBSERVICES CRIADA EM 04/07/2019
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







Function requestAPI(url, str, way)
	
	strWS = replace(str, ":", "=")
	strWS = replace(strWS, "/", "&")
	
	way = UCase(way)
	strWS = LCase(strWS)
	url = LCase(url)
	
	mac = "teste"
	
	strRequest="{""mac"":""" & mac & """}"
	
	if way = "GET" then
	
		EndPointLink = url & "?" & strWS
	
	else
	
		EndPointLink = url & "" & strWS
	
	end if
	
	
	dim http
	set http=createObject("Microsoft.XMLHTTP")
	http.open "GET",EndPointLink,false
	http.setRequestHeader "Content-Type","application/json"
	http.setRequestHeader "X-Parse-Application-Id","XXXXXXXXXXXXXXXXXXXXX"

	' msgbox "REQUEST : " & strRequest
	http.send strRequest
	
	If http.Status = 200 Then
		' msgbox "RESPONSE : " & http.responseText
		responseText = http.responseText
	
		decode1 = json_decode(responseText)
		
	else
		' msgbox "ERRCODE : " & http.status
		decode1 = "erro"
	End If
	
	requestAPI = decode1
End Function

Function json_decode(arr)
	dim arrayGiga(4000)
	limitador = 0
	decode1 = arr
	decode1 = Replace(decode1, "{", "")
	decode1 = Replace(decode1, "}", "")
	decode1 = Replace(decode1, """", "")
	decode1 = Replace(decode1, """", "")
	decode1 = Replace(decode1, "", " ")
	decode1 = TRIM(decode1)
	
	contador = 0
	' Usar dicionário!!!!!!
	
	valores = split(decode1, ",")
	
	while contador <= UBound(valores)
	
		chave_valor = split(valores(contador), ":")
		
		meConta = 0
		
		while meConta <= UBound(chave_valor)
			
			arrayGiga(limitador) = chave_valor(meConta)
			limitador = limitador + 1
			
			meConta = meConta + 1
		wend
	
		contador = contador + 1
	
	wend
	
	
	
	'Testando arrayGiga
	Bond = 0
	while Bond < limitador
		' msgbox arrayGiga(Bond)
		
		Bond = Bond + 1
	wend
	
	arrayGiga(4000) = limitador
	
		'decodeFINAL = Split(decode1, " ")
		
		
		'fff = 1
		
		'while fff <= UBound(decodeFINAL)
		'	decode1 = decode1 & " " & decodeFINAL(fff)
		'	fff = fff + 1
		'wend
	
	json_decode = arrayGiga

End Function

'Variavel = RequestAPI("http://localhost/vbstest.php", "moeda:dolar", "GET") 'O certo aqui é pimbe
'Limitador = Variavel(4000)

'a = 0
'while a < Limitador
'	msgbox Variavel(a) & " => " & Variavel(a+1)
'	a = a + 2
'wend

