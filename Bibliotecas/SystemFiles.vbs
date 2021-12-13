'        @PARANAUERJ DEVELOPEMENT op
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







set FSO = CreateObject("Scripting.FileSystemObject")
set WShell = WScript.CreateObject("WScript.Shell")
set Comando = WScript.CreateObject("WScript.Shell")
Set utf8 = CreateObject("System.Text.UTF8Encoding")
Set b64Enc = CreateObject("System.Security.Cryptography.ToBase64Transform")
Set b64Dec = CreateObject("System.Security.Cryptography.FromBase64Transform")
Set mac = CreateObject("System.Security.Cryptography.HMACSHA256")
Set aes = CreateObject("System.Security.Cryptography.RijndaelManaged")
Set mem = CreateObject("System.IO.MemoryStream")


Function Mover(a,b)
	on error resume next
	if FSO.FileExists(a) then
		FSO.MoveFile a,b
		Mover = "Arquivo movido com sucesso!"
	elseif FSO.FolderExists(a) then
		FSO.MoveFolder a,b
		Mover = "Pasta movida com sucesso!"
	else
		Mover = "Nao foi possivel completar a acao!"
	end if
	end function


Function Move(a,b)
	on error resume next
	if FSO.FileExists(a) then
		FSO.MoveFile a,b
		Mover = "Arquivo movido com sucesso!"
	elseif FSO.FolderExists(a) then
		FSO.MoveFolder a,b
		Mover = "Pasta movida com sucesso!"
	else
		Mover = "Nao foi possivel completar a acao!"
	end if
	end function


Function Copiar(a,b)
	on error resume next
	if FSO.FileExists(a) then
		FSO.CopyFile a,b
		Copiar = "Arquivo copiado com sucesso!"
	elseif FSO.FolderExists(a) then
		Copiar = "Pasta copiada com sucesso!"
	else
		msgbox"Nao foi possivel completar a acao!",0,"Erro"
	end if
	end function

Function Copy(a,b)
	on error resume next
	if FSO.FileExists(a) then
		FSO.CopyFile a,b
		Copiar = "Arquivo copiado com sucesso!"
	elseif FSO.FolderExists(a) then
		Copiar = "Pasta copiada com sucesso!"
	else
		msgbox"Nao foi possivel completar a acao!",0,"Erro"
	end if
	end function
	
	
Function Deletar(a)
	on error resume next
	if FSO.FileExists(a) then
		FSO.DeleteFile(a)
		Deletar = "Arquivo deletado com sucesso!"
	elseif FSO.FolderExists(a) then
		FSO.DeleteFolder(a)
		Deletar = "Pasta deletada com sucesso!"
	else
		msgbox"Nao foi possivel completar a acao!",0,"Erro"
	end if
	end function


Function Delete(a)
	on error resume next
	if FSO.FileExists(a) then
		FSO.DeleteFile(a)
		Deletar = "Arquivo deletado com sucesso!"
	elseif FSO.FolderExists(a) then
		FSO.DeleteFolder(a)
		Deletar = "Pasta deletada com sucesso!"
	else
		msgbox"Nao foi possivel completar a acao!",0,"Erro"
	end if
	end function

	
Function Sair()
	msgbox"Encerrando terminal..."
	WScript.Quit
	end function

Function ExitVPP()
	msgbox"Encerrando terminal..."
	WScript.Quit
	end function
	

Function Pesquisar(a)
	b = Replace(a, "_", "%20")
	WShell.run "microsoft-edge:" & b
	end function
	

Function Search(a)
	b = Replace(a, "_", "%20")
	WShell.run "microsoft-edge:" & b
	end function
	
Function Ping(a,b)
	WShell.run"ping " & a & " -n " & b
	end function
	
	
Function Traceroute(a)
	WShell.run"tracert " & a
	end function
	
	 
Function Reiniciar()
	WShell.run "main.wsf"
	WScript.Quit
	end function


Function RestartPC()
	WShell.run "main.wsf"
	WScript.Quit
	end function
	
	
Function matarTask(a)
	WShell.run"cmd /k taskkil /f /im " & a
	end function


Function tryUpdate()
	set arqConf = FSO.OpenTextFile(Comando.CurrentDirectory & "\Config\Config.conf", 1)
	linhaCont = 0
	Do Until arqConf.AtEndOfStream
		linha = arqConf.Readline
			if linhaCont = 4 then
				linhaSplitada = split(linha)
				'msgbox linhaSplitada(1)
				Dim o
				Set o = CreateObject("MSXML2.XMLHTTP")
				o.open "GET", "https://vbsendpoints.000webhostapp.com/versionUpdate.php?version=" & linhaSplitada(1), False
				o.send
				if o.Status = 200 then
					att = o.responseText
					if att = "atualizado" then
						msgbox "Seu compilador ja esta atualizado!"
					else
						WShell.run "Updater.vbs"
						WScript.Quit
					end if
				else 
					msgbox "Erro de conexao"
				end if
				
			end if
		linhaCont = linhaCont + 1
		Loop
	arqConf.Close
	end function
	

Function addExtension(a)
	msgbox "Adicionando extensao " & a & "..."
	dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
	dim bStrm: Set bStrm = createobject("Adodb.Stream")
	xHttp.Open "GET", "https://vbsendpoints.000webhostapp.com/Extensions/" & a & ".vbs", False
	xHttp.Send
	if xHttp.Status = 200 then
		with bStrm
			.type = 1 '//binary
			.open
			.write xHttp.responseBody
			.savetofile Comando.CurrentDirectory & "\Extensoes\" & UCase(a) & ".vbs", 2 '//overwrite
		end with
		bStrm.Close

		Const ForReading = 1
		Const ForWriting = 2
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.OpenTextFile(Comando.CurrentDirectory & "\main.wsf", ForReading)
		strText = objFile.ReadAll
		objFile.Close

		strNewText = Replace(strText, "<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Extensoes/"& a &".vbs"&chr(34)&"/>", "")
		strNewText = Replace(strNewText, "Break", "<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Extensoes/"& a &".vbs"&chr(34)&"/>" & vbCrlf & "Break")

		strNewText = TRIM(strNewText)
		Set objFile = objFSO.OpenTextFile(Comando.CurrentDirectory & "\main.wsf", ForWriting)
		objFile.Write strNewText
		objFile.Close

		WShell.run "main.wsf"
		WScript.Quit

	else
		msgbox "Nao foi possivel adicionar esta biblioteca!"
	end if
end function


Function exibeMensagem(a)
	if a = 1 then 
		msgbox"Comando invalido!",0,"Erro"
	elseif a = 2 then
		msgbox"Sintaxe invalida ou arquivo/diretorio nao encontrado!" ,0,"Erro"
	end if
	end function 
	
	
Function Maquina(a)
	if a = "TURNOFF" then
		WShell.run"cmd /k shutdown -s /f /c Desligando..."
		end if
	if a = "RESTART" then
		WShell.run"cmd /k shutdown -r /f /c Reiniciando..."
		end if
	if a = "HIBERNATE" then
		WShell.run"cmd /k shutdown -H"
		end if
	if a = "LOGOFF" then
		WShell.run"cmd /k shutdown -L"
		end if
		end function
	
	
Function Programar()
	msgbox"Funcionalidade disponível no comando Compilar"
	end function
	
	
Function Abrir(a)
	on error resume next
	WShell.run "" & a
	end function
	
Function AbrirProjeto(a, b, c)
	if FSO.FileExists ("Projetos\" & a & "\" & a & "_COMPILADO.wsf") then
		if b = "/CONSOLE" then
				Comando.run "Cscript Projetos\" & a & "\" & a & "_COMPILADO.wsf " & c
			elseif b = "/WINDOW" then
				Comando.run "WScript Projetos\" & a & "\" & a & "_COMPILADO.wsf " & c
			else
				Comando.run "Wscript Projetos\" & a & "\" & a & "_COMPILADO.wsf " & c
			end if
	else
		msgbox "Arquivo nao localizado: " & "Projetos\" & a & "\" & a & "_COMPILADO.wsf"
	end if
	end function
	
Function EditarProjeto(a)
	if FSO.FileExists ("Projetos\" & a & "\" & a & ".vpp") then
		WShell.run "notepad++ " & Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & ".vpp"
	else
		msgbox "Arquivo nao encontrado: " & "Projetos\" & a & "\" & a & ".vpp"
	end if
	end function

Function Rede()
	WShell.run"cmd /k ipconfig /all"
	end function


Function Esperar(a)
	ConstantSeconds = a * 1000
	Wscript.Sleep ConstantSeconds
	end function
	
Function Wait(a)
	ConstantSeconds = a * 1000
	Wscript.Sleep ConstantSeconds
	end function

Function ExportarProjeto(a,b)
	if NOT FSO.FolderExists(b) then
		FSO.CreateFolder(b)
	end if
			if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & "_COMPILADO.wsf") then
				FSO.CopyFolder Comando.CurrentDirectory & "\Projetos\" & a, b
			else
				msgbox "Projeto inexistente: " & a
				end if
	end function


Function ImportarProjeto(a,b)
	if NOT FSO.FolderExists(b) then
			msgbox "Diretorio " & b & " nao existe"
		else
		if FSO.FolderExists (Comando.CurrentDirectory & "\Projetos\" & a & "") then
			msgbox "Projeto com mesmo nome ja existente: " & a
		else
			FSO.CopyFolder b, Comando.CurrentDirectory & "\Projetos\"
		end if
		end if	
	end function

	
Function CriarArquivo(a)
	msgbox"Funciona assim: voce vai escrever o nome do arquivo (sem extensao), altera ele (pode ser com o notepad mesmo) e compila com o comando COMPILA"
	FSO.CreateFolder(Comando.CurrentDirectory & "\Projetos\" & a & "")
	Wscript.Sleep 500
	FSO.CreateFolder(Comando.CurrentDirectory & "\Projetos\" & a & "\XML")
	FSO.CreateFolder(Comando.CurrentDirectory & "\Projetos\" & a & "\Bibliotecas")
	FSO.CreateFolder(Comando.CurrentDirectory & "\Projetos\" & a & "\Encoded")
	FSO.CreateFolder(Comando.CurrentDirectory & "\Projetos\" & a & "\Extensoes")
	FSO.CreateFolder(Comando.CurrentDirectory & "\Projetos\" & a & "\Classes")
	FSO.CreateFolder(Comando.CurrentDirectory & "\Projetos\" & a & "\Media")
	set arq = FSO.CreateTextFile(Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & ".vpp", true)
		arq.WriteLine("Import Lib.Math")
		arq.WriteLine("Import Lib.SystemFiles")
		arq.WriteLine("")
		arq.WriteLine("")
		arq.WriteLine("")
		arq.WriteLine("")
		arq.WriteLine("")
		arq.WriteLine("")
		arq.WriteLine("")
		arq.WriteLine("Start Code Nome_do_Projeto")
		arq.WriteLine("")
		arq.WriteLine("Var Statement 1")
		arq.WriteLine("	Var Integer i1")
		arq.WriteLine("")
		arq.WriteLine("Main")
		arq.WriteLine("")
		arq.WriteLine("// O codigo principal fica aqui!")
		arq.WriteLine("")
		arq.WriteLine("End Code")
		arq.Close
	WShell.run "notepad++ " & Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & ".vpp"
	end function
	
	
Function Projetos()
	pasta = Comando.CurrentDirectory & "\Projetos"
	projetos = ""
	For each arquivo in FSO.GetFolder(pasta).SubFolders
		projetos = projetos & arquivo & vbCrlf
	Next
	projetos = Replace(projetos, Comando.CurrentDirectory & "\Projetos\", " ")
	Msgbox projetos
	
	end function
	


Function ucFirst(str)
	str = LCase(str)
	str = UCase(Left(str, 1)) &  Mid(str, 2)

	ucFirst = str
	
	end function



Function getText(file)
	Set arq = FSO.OpenTextFile(file, 1)	
	text = ""
	i = 0
	Do Until arq.AtEndOfStream
		if i > 0 then
			text = vbCrlf & text
		end if

		text = text & arq.ReadLine
		
		i = i + 1
	Loop

	arq.Close

	getText = text

	end function


Function getTextLines(file)
	Set arq = FSO.OpenTextFile(file, 1)	

	Dim objDictionary
	Set objDictionary = CreateObject("Scripting.Dictionary")
	objDictionary.CompareMode = vbTextCompare

	i = 0
	Do Until arq.AtEndOfStream

		objDictionary.Add i, arq.ReadLine
		
		i = i + 1
	Loop

	arq.Close

	SET getTextLines = objDictionary

	end function


Function getNLines(file)
	Set arq = FSO.OpenTextFile(file, 1)	

	i = 0
	Do Until arq.AtEndOfStream
		arq.readLine()
		i = i + 1
	Loop

	arq.Close

	getNLines = i

	end function


Function WriteTextLine(file, str)
	Set arq = FSO.OpenTextFile(file, 8)	

	arq.WriteLine(str)

	arq.Close

	end function

Function WriteText(file, str)
	Set arq = FSO.OpenTextFile(file, 8)	

	arq.Write(str)

	arq.Close

	end function

Function setText(file, str)
	Set arq = FSO.OpenTextFile(file, 2)	

	arq.Write(str)

	arq.Close

	end function


Function clearFile(file)
	Set arq = FSO.OpenTextFile(file, 2)	

	arq.Write("")

	arq.Close

	end function


Function Extensoes()
	pasta = Comando.CurrentDirectory & "\Extensoes"
	ext = ""
	Set objFolder = FSO.GetFolder(pasta)
	Set colFiles = objFolder.Files

	For each arquivo in colFiles
		ext = ext & arquivo & vbCrlf
	Next
	ext = Replace(ext, Comando.CurrentDirectory & "\Extensoes\", " ")
	Msgbox ext
	
	end function



Function isset(val)

    isset2 = false
    if IsNull(val) or val = "" then 
		isset2 = false
	else
		isset2 = true
	end if

	isset = isset2

End Function


Function convert(val, tipo)

	tipoAux = UCase(tipo)

	res = 0

	select case tipoAux

		case "INTEGER"
			res = CInt(val)
		case "STRING"
			res = CStr(val)
		case "FLOAT"
			res = CDbl(val)
		case "DOUBLE"
			res = CDbl(val)
		case "BOOLEAN"
			res = CBool(val)
		case "INT"
			res = CInt(val)
		case "DATE"
			res = CDate(val)
		case "STR"
			res = CStr(val)
		case "BOOL"
			res = CBool(val)
		case else
			res = val

	End select

	convert = res

End Function


Function VType(variavel)

	tipo = varType(variavel)

	select case tipo

		case 0
			res = "Empty"
		case 1
			res = "Null"
		case 2
			res = "Integer"
		case 3
			res = "Long"
		case 4
			res = "Single"
		case 5
			res = "Double"
		case 6
			res = "Currency"
		case 7
			res = "Date"
		case 8
			res = "String"
		case 9 
			res = "Object"
		case 10
			res = "Error"
		case 11
			res = "Boolean"
		case 12
			res = "Variant"
		case 13
			res = "Data-object"
		case 17
			res = "Byte"
		case 8192
			res = "Array"
		case else
			res = "XXX"

	End select

	VType = UCase(res)

End Function



Function random(maximo, minimo)
	rand = 0
	Randomize
	rand = Int((maximo-minimo+1)*Rnd+minimo)
	random = rand
end Function



Function vppDate(param1)

	a = now()

	if isset(param1) then
		if UCase(param1) = "DAY" then
			retorno = day(a)
		elseif UCase(param1) = "MONTH" then
			retorno = month(a)
		elseif UCase(param1) = "MONTHNAME" then
			retorno = MonthName(month(a))
		elseif UCase(param1) = "WEEKDAY" then
			retorno = WeekDay(a)
		elseif UCase(param1) = "WEEKDAYNAME" then
			retorno = WeekDayName(a)
		elseif UCase(param1) = "YEAR" then
			retorno = year(a)
		elseif UCase(param1) = "ALL" then
			retorno = date()
		else
			msgbox "Parametro errado na funcao today()!"
		end if
	else
		retorno = date()
	end if
	
	vppDate = UCase(retorno)

End Function


Function vppTime(param1)

	a = now()

	if isset(param1) then
		if UCase(param1) = "HOUR" then
			retorno = hour(a)
		elseif UCase(param1) = "MINUTE" then
			retorno = minute(a)
		elseif UCase(param1) = "SECOND" then
			retorno = second(a)
		elseif UCase(param1) = "ALL" then
			retorno = time()
		else
			msgbox "Parametro errado na funcao today()!"
		end if
	else
		retorno = time()
	end if
	
	vppTime = UCase(retorno)

End Function

'Function RebuildMain()
'
'	set Comando = WScript.CreateObject("WScript.Shell")
'	Set arq = FSO.CreateTextFile(Comando.CurrentDirectory & "Main.wsf", true)
'	arq.WriteLine("<job id="&chr(34)&"compilador-com-vbs"&chr(34)&">")
'	
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&" src="&chr(34)&"Config/Recursos/VerVersaoVBS.vbs"&chr(34)&"/>")
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&" src="&chr(34)&"Config/Recursos/XMLProjectGenerator.vbs"&chr(34)&"/>")
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&" src="&chr(34)&"Bibliotecas/Math.vbs"&chr(34)&"/>")
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&" src="&chr(34)&"Bibliotecas/SystemFiles.vbs"&chr(34)&"/>")
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&" src="&chr(34)&"Bibliotecas/Comunicacao.vbs"&chr(34)&"/>")
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&" src="&chr(34)&"Bibliotecas/BD.vbs"&chr(34)&"/>")
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&" src="&chr(34)&"Extensoes/FUNCOES.vbs"&chr(34)&"/>")
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&" src="&chr(34)&"Controller/controlador.vbs"&chr(34)&"/>")
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&" src="&chr(34)&"Interface/terminal.vbs"&chr(34)&"/>")
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&" src="&chr(34)&"Interpretador/interpretadorbasico.vbs"&chr(34)&"/>")
'	arq.WriteLine("")
'	arq.WriteLine("nhau")
'	arq.WriteLine("<script language="&chr(34)&"VBScript"&chr(34)&">")
'	arq.WriteLine("rodarTerminal()")
'	arq.WriteLine("</script>")
'	arq.WriteLine("</job>")
'
'	end function


Function Input(myPrompt)
    ' Check if the script runs in CSCRIPT.EXE
    If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
        ' If so, use StdIn and StdOut
		'WScript.Echo vbCrlf
        WScript.StdOut.Write myPrompt & " "
        Input = WScript.StdIn.ReadLine
    Else
        ' If not, use InputBox( )
		res = InputBox(myPrompt)
        Input = res
    End If
End Function



Function Pause(strPause)
     ' Check if the script runs in CSCRIPT.EXE
    If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
		WScript.Echo (vbCrlf & vbCrlf & "Press Enter to continue")
		z = WScript.StdIn.ReadLine()
	End if
End Function



Function Message(str)
	
	strOutput = replaceUTF(str)

	If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
		WScript.Echo strOutput
        'WScript.Echo vbcrlf & str
    Else
        msgbox strOutput,,"Vpp"
    End If
End Function


Function callRequireds(arr)

	cont = 0

	For Each obj in arr.keys

		if arr(obj)(1) = "TRUE" then
			If Not( WScript.Arguments.Named.Exists(arr(obj)(0)) ) Then
				WScript.Arguments.ShowUsage
				WScript.Quit
			End If
		end if

		cont = cont + 1

	next

end function


Function getArg(val)

	if WScript.Arguments.Named.Exists(val) then
		retorno = WScript.Arguments.Named(val)
	else
		retorno = "NULL"
	end if

	getArg = retorno

end function


function UFirst(str)

	b = Left(str, 1)
	str = replace(str, b, UCase(b), 1, 2)
	UFirst = str

end function

function readBool(str)

	answer = Comando.Popup(str, infinite, "VPP Compiler", vbQuestion + vbYesNo + vbDefaultButton2)

	If answer = vbYes Then
		retorno = true
	Else
		retorno = false
	End If

	readBool = retorno

end function


function LPad(str, preenchedor, tamanho)
	
	retorno = Right(preenchedor & str, tamanho)

	lpad = retorno

end function


Function Min(a, b)
    Min = a
    If b < a Then Min = b
End Function


' Convert a byte array to a Base64 string representation of it.
'
' Arguments:
'   bytes (Byte()): Byte array.
'
' Returns:
'   String: Base64 representation of the input byte array.
Function B64Encode(bytes)
    blockSize = b64Enc.InputBlockSize
    For offset = 0 To LenB(bytes) - 1 Step blockSize
        length = Min(blockSize, LenB(bytes) - offset)
        b64Block = b64Enc.TransformFinalBlock((bytes), offset, length)
        result = result & utf8.GetString((b64Block))
    Next
    B64Encode = result
End Function


' Convert a Base64 string to a byte array.
'
' Arguments:
'   b64Str (String): Base64 string.
'
' Returns:
'   Byte(): A byte array that the Base64 string decodes to.
Function B64Decode(b64Str)
    bytes = utf8.GetBytes_4(b64Str)
    B64Decode = b64Dec.TransformFinalBlock((bytes), 0, LenB(bytes))
End Function


' Concatenate two byte arrays.
'
' Arguments:
'   a (Byte()): A byte array.
'   b (Byte()): Another byte array.
'
' Returns:
'   Byte(): Concatenated byte arrays.
Function ConcatBytes(a, b)
    mem.SetLength(0)
    mem.Write (a), 0, LenB(a)
    mem.Write (b), 0, LenB(b)
    ConcatBytes = mem.ToArray()
End Function


' Check if two byte arrays are equal.
'
' Arguments:
'   a (Byte()): A byte array.
'   b (Byte()): Another byte array.
'
' Returns:
'   Boolean: True if both byte arrays are equal; False otherwise.
Function EqualBytes(a, b)
    EqualBytes = False
    If LenB(a) <> LenB(b) Then Exit Function
    diff = 0
    For i = 1 to LenB(a)
        diff = diff Or (AscB(MidB(a, i, 1)) Xor AscB(MidB(b, i, 1)))
    Next
    EqualBytes = Not diff
End Function


' Compute message authentication code using HMAC-SHA-256.
'
' Arguments:
'   msgBytes (Byte()): Message to be authenticated.
'   keyBytes (Byte()): Secret key.
'
' Returns:
'   Byte(): Message authenticate code.
Function ComputeMAC(msgBytes, keyBytes)
    mac.Key = keyBytes
    ComputeMAC = mac.ComputeHash_2((msgBytes))
End Function


' Encrypt plaintext and compute MAC for the result.
'
' The length of AES encryption key (aesKey) must be 256 bits (32 bytes).
' It must be provided as a Base64 encoded string. On macOS or Linux,
' enter this command to generate a Base64 encoded 256-bit key:
'
'   head -c32 /dev/urandom | base64
'
' The HMAC secret key (macKey) can be any length but a minimum of
' 256 bits (32 bytes) is recommended as the length of this key. It must
' be provided as a Base64 encoded string.
'
' The return value of this function is composed of the following three
' Base64 encoded strings joined with colons:
'
'   - Message authentication code.
'   - Randomly generated 128-bit initialization vector (IV).
'   - Ciphertext.
'
' Note:
'
'   - A 256-bit key after Base64 encoding contains 44 characters
'     including one '=' character as padding at the end.
'   - A 128-bit IV after Base64 encoding contains 24 characters
'     including two '=' characters as padding at the end.
'
' Arguments:
'   plaintext (String): Text to be encrypted.
'   aesKey (String): AES encryption key encoded as a Base64 string.
'   macKey (String): HMAC secret key encoded as a Base64 string.
'
' Returns:
'   String: MAC, IV, and ciphertext joined with colons.
Function Encrypt(plaintext, aesKey, macKey)
    aes.GenerateIV()
    aesKeyBytes = B64Decode(aesKey)
    macKeyBytes = B64Decode(macKey)
    Set aesEnc = aes.CreateEncryptor_2((aesKeyBytes), aes.IV)
    plainBytes = utf8.GetBytes_4(plaintext)
    cipherBytes = aesEnc.TransformFinalBlock((plainBytes), 0, LenB(plainBytes))
    macBytes = ComputeMAC(ConcatBytes(aes.IV, cipherBytes), macKeyBytes)
    Encrypt = B64Encode(macBytes) & ":" & B64Encode(aes.IV) & ":" & _
              B64Encode(cipherBytes)
End Function


' Decrypt ciphertext after authenticating IV and ciphertext using MAC.
'
' MAC, IV, and ciphertext must be encoded in Base64. They are provided
' together as a single string with the Base64 encoded values separated
' by colons. See the comment for Encrypt() function to read more about
' the format.
'
' Arguments:
'   macIVCipherText (String): Colon separated MAC, IV, and ciphertext.
'   aesKey (String): AES encryption key encoded as a Base64 string.
'   macKey (String): HMAC secret key encoded as a Base64 string.
'
' Returns:
'   String: Plaintext that the given ciphertext decrypts to.
Function Decrypt(macIVCiphertext, aesKey, macKey)
    aesKeyBytes = B64Decode(aesKey)
    macKeyBytes = B64Decode(macKey)
    tokens = Split(macIVCiphertext, ":")
    macBytes = B64Decode(tokens(0))
    ivBytes = B64Decode(tokens(1))
    cipherBytes = B64Decode(tokens(2))
    macActual = ComputeMAC(ConcatBytes(ivBytes, cipherBytes), macKeyBytes)
    If Not EqualBytes(macBytes, macActual) Then
        Err.Raise vbObjectError + 1000, "Decrypt()", "Bad MAC"
    End If
    Set aesDec = aes.CreateDecryptor_2((aesKeyBytes), (ivBytes))
    plainBytes = aesDec.TransformFinalBlock((cipherBytes), 0, LenB(cipherBytes))
    Decrypt = utf8.GetString((plainBytes))
End Function


function replaceUTF(str)
	
	retorno = str
	
	'Maiusculas
	retorno = replace(retorno, "À", chr(192))
	retorno = replace(retorno, "Á", chr(193))
	retorno = replace(retorno, "Â", chr(194))
	retorno = replace(retorno, "Ã", chr(195))
	retorno = replace(retorno, "È", chr(200))
	retorno = replace(retorno, "É", chr(201))
	retorno = replace(retorno, "Ê", chr(202))
	retorno = replace(retorno, "Ì", chr(204))
	retorno = replace(retorno, "Í", chr(205))
	retorno = replace(retorno, "Ò", chr(210))
	retorno = replace(retorno, "Ó", chr(211))
	retorno = replace(retorno, "Ô", chr(212))
	retorno = replace(retorno, "Õ", chr(213))
	retorno = replace(retorno, "Ú", chr(218))
	retorno = replace(retorno, "Ç", chr(199))
	
	'Minusculas
	retorno = replace(retorno, "à", chr(224))
	retorno = replace(retorno, "á", chr(225))
	retorno = replace(retorno, "â", chr(226))
	retorno = replace(retorno, "ã", chr(227))
	retorno = replace(retorno, "è", chr(232))
	retorno = replace(retorno, "é", chr(233))
	retorno = replace(retorno, "ê", chr(234))
	retorno = replace(retorno, "ì", chr(236))
	retorno = replace(retorno, "í", chr(237))
	retorno = replace(retorno, "ò", chr(242))
	retorno = replace(retorno, "ó", chr(243))
	retorno = replace(retorno, "ô", chr(244))
	retorno = replace(retorno, "õ", chr(245))
	retorno = replace(retorno, "ú", chr(250))
	retorno = replace(retorno, "ç", chr(231))
	
	replaceUTF = retorno

end function