'        @PARANAUERJ DEVELOPEMENT WITH UPDATE
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







set FSO = CreateObject("Scripting.FileSystemObject")



Function Compila(a, codifica)
	dim purinho
	purinho = a
	set Comando = WScript.CreateObject("WScript.Shell")
	Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & ".vpp")	
	WScript.Sleep 700
	Dim linha (10000)
	y = 0
	LOOPVAR = 0
	Do Until arq.AtEndOfStream
		linha(y) = arq.ReadLine
		' linha(y) = UCase(linha(y))
		linha(y) = limpaLinha(linha(y))
		linha(y) = TRIM(linha(y))
		y = y + 1
		Loop
		
	totallinhas = arq.Line-1
	arq.Close
	errolib = 0
	
	
	set arq2 = FSO.CreateTextFile(Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & "_COMPILADO.wsf", true)
	linha5 = Split(linha(9))
	if UBound(linha5) <> 2 then
			Message"Erro ao iniciar o codigo." & vbCrlf & "Linha 10"
			Exit Function
Compila = false
		elseif UBound(linha5) = 2 then
		if UCase(linha5(0)) = "START" and UCase(linha5(1)) = "CODE" then
			arq2.WriteLine "<job id="& chr(34) & linha5(2) & chr(34) &">"
		else
			Message"Erro ao iniciar o codigo." & vbCrlf & "Linha 10"
			Exit Function
Compila = false
		end if
		else 
			Message"Erro ao iniciar o codigo." & vbCrlf & "Linha 10"
			Exit Function
Compila = false
		end if
	x = 0
	linha6 = Split(linha(10))
	
	if UBound(linha6) >= 0 then
		Message"Caracteres nao sao permitidos na linha." & vbcrlf & "Linha 11"
		Exit Function
Compila = false
	end if
	
	libmath = 0
	libsystemfiles = 0
	libbd = 0
	libmedia = 0
	libarray = 0
	libjson = 0

	nArgs = 0
	Set args = CreateObject("Scripting.Dictionary")
	'Set argsAux = CreateObject("System.Collections.ArrayList")

	while x < 9
		linhasplit = Split(linha(x))
		if UCase(linha(x)) = "IMPORT LIB.MATH" then
			arq2.WriteLine"<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Bibliotecas/Math.vbs"&chr(34)&"/>"
			FSO.CopyFile Comando.CurrentDirectory & "\Bibliotecas\Math.vbs", Comando.CurrentDirectory & "\Projetos\" & a & "\Bibliotecas\"
			libmath = 1
			
		elseif UCase(linha(x)) = "IMPORT LIB.SYSTEMFILES" then
			arq2.WriteLine"<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Bibliotecas/SystemFiles.vbs"&chr(34)&"/>"
			FSO.CopyFile Comando.CurrentDirectory & "\Bibliotecas\SystemFiles.vbs", Comando.CurrentDirectory & "\Projetos\" & a & "\Bibliotecas\"
			libsystemfiles = 1

		elseif UCase(linha(x)) = "IMPORT LIB.WEB" then
			arq2.WriteLine"<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Bibliotecas/Web.vbs"&chr(34)&"/>"
			FSO.CopyFile Comando.CurrentDirectory & "\Bibliotecas\Web.vbs", Comando.CurrentDirectory & "\Projetos\" & a & "\Bibliotecas\"

		elseif UCase(linha(x)) = "IMPORT LIB.JSON" then
			arq2.WriteLine"<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Bibliotecas/JSON.vbs"&chr(34)&"/>"
			FSO.CopyFile Comando.CurrentDirectory & "\Bibliotecas\JSON.vbs", Comando.CurrentDirectory & "\Projetos\" & a & "\Bibliotecas\"
			arq2.WriteLine"<script language=" & chr(34) & "JScript" & chr(34) &" src="&chr(34) &"Bibliotecas/JSON.js"&chr(34)&"/>"
			FSO.CopyFile Comando.CurrentDirectory & "\Bibliotecas\JSON.js", Comando.CurrentDirectory & "\Projetos\" & a & "\Bibliotecas\"
			libjson = 1

		elseif UCase(linha(x)) = "IMPORT BOOTSTRAP" then
			FSO.CopyFolder Comando.CurrentDirectory & "\Bibliotecas\Bootstrap", Comando.CurrentDirectory & "\Projetos\" & a & "\Bibliotecas\Bootstrap"
			
		elseif UCase(linha(x)) = "IMPORT LIB.MEDIA" then
			arq2.WriteLine"<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Bibliotecas/Media.vbs"&chr(34)&"/>"
			FSO.CopyFile Comando.CurrentDirectory & "\Bibliotecas\Media.vbs", Comando.CurrentDirectory & "\Projetos\" & a & "\Bibliotecas\"
			libmedia = 1

		elseif UCase(linha(x)) = "IMPORT LIB.ARRAY" then
			arq2.WriteLine"<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Bibliotecas/Array.vbs"&chr(34)&"/>"
			FSO.CopyFile Comando.CurrentDirectory & "\Bibliotecas\Array.vbs", Comando.CurrentDirectory & "\Projetos\" & a & "\Bibliotecas\"
			libarray = 1
		
		elseif UCase(linha(x)) = "IMPORT LIB.BD" then
			arq2.WriteLine"<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Bibliotecas/BD.vbs"&chr(34)&"/>"
			if NOT FSO.FolderExists (Comando.CurrentDirectory & "\Projetos\" & a & "\Database") then
			FSO.CreateFolder Comando.CurrentDirectory & "\Projetos\" & a & "\Database"
			end if
			FSO.CopyFile Comando.CurrentDirectory & "\Bibliotecas\BD.vbs", Comando.CurrentDirectory & "\Projetos\" & a & "\Bibliotecas\"
			libbd = 1
			
			
		elseif UBound(linhasplit) >= 1 then
			if UCase(linhasplit(0)) = "INCLUDE" then
				if UBound(linhasplit) = 1 then
					if FSO.FileExists(Comando.CurrentDirectory & "\Extensoes\" & UCase(linhasplit(1))) then
						arq2.WriteLine"<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Extensoes/" & UCase(linhasplit(1)) & chr(34)&"/>"
						FSO.CopyFile Comando.CurrentDirectory & "\Extensoes\" & UCase(linhasplit(1)), Comando.CurrentDirectory & "\Projetos\" & a & "\Extensoes\"
					else
						Message "Arquivo nao encontrado: " & Comando.CurrentDirectory & "\Extensoes\" & UCase(linhasplit(1)) & vbCrlf & "Linha: " & x+1
						Exit Function
Compila = false
					end if
				else
					Message "E necessario escolher apenas um arquivo apos o include!" & vbCrlf & "Linha: " & x+1
					Exit Function
Compila = false
				end if
				else
				if UCase(linhasplit(0)) = "CINCLUDE" then
				if UBound(linhasplit) = 1 then
					if FSO.FileExists(Comando.CurrentDirectory & "\Classes\" & UCase(linhasplit(1)) & ".class.vpp") then
						generateClass(UCase(linhasplit(1)))
						WScript.Sleep 400
						arq2.WriteLine"<script language=" & chr(34) & "VBScript" & chr(34) &" src="&chr(34) &"Classes/" & UCase(linhasplit(1)) & ".class.vbs" & chr(34)&"/>"
						FSO.CopyFile Comando.CurrentDirectory & "\Classes\Generated\" & UCase(linhasplit(1)) & ".class.vbs", Comando.CurrentDirectory & "\Projetos\" & a & "\Classes\"
						FSO.CopyFile Comando.CurrentDirectory & "\Classes\" & UCase(linhasplit(1)) & ".class.vpp", Comando.CurrentDirectory & "\Projetos\" & a & "\Classes\"
					else
						Message "Arquivo nao encontrado: " & Comando.CurrentDirectory & "\Classes\" & UCase(linhasplit(1)) & ".class.vpp" & vbCrlf & "Linha: " & x+1
						Exit Function
Compila = false
					end if
				else
					Message "E necessario escolher apenas um arquivo apos o cinclude!" & vbCrlf & "Linha: " & x+1
					Exit Function
Compila = false
				end if
				else
				if UBound(linhasplit) >= 1 then
				
				if UCase(linhasplit(0)) = "SYSTEM" then
					argse = Replace(linha(x), "SYSTEM ", "")
					argse = Replace(linha(x), "System ", "")
					argse = Replace(linha(x), "system ", "")

					argsAux = Split(argse, "(", 2)
					argsAux(0) = TRIM(argsAux(0))

					if  UCase(argsAux(0)) = "SYSTEM ARGS" then
						argsAux(1) = Replace(argsAux(1), "(", "")
						argsAux(1) = Replace(argsAux(1), ")", "")
						argumentos = Split(argsAux(1), ";")
						contArgs = 0
						arq2.WriteLine("<runtime>")
						while contArgs <= UBound(argumentos)

							argumentos(contArgs) = Replace(argumentos(contArgs), "[", "")
							argumentos(contArgs) = Replace(argumentos(contArgs), "]", "")
							argumentos(contArgs) = Replace(argumentos(contArgs), chr(34), "")

							elArgs = split(argumentos(contArgs), ",")
							elArgs(0) = TRIM(elArgs(0))
							elArgs(1) = TRIM(elArgs(1))
							elArgs(2) = TRIM(elArgs(2))

							if UBound(elArgs) < 2 then
								Message "Numero de parametros invalidos para um argumento do sistema!" & vbCrlf & "Linha: " & x+1
								Exit Function
Compila = false
							end if
							if UCase(elArgs(2)) <> "TRUE" and UCase(elArgs(2)) <> "FALSE" then
								msgbox elArgs(2)
								Message "O Terceiro parametro do argumento deve ser true ou false!" & vbCrlf & "Linha: " & x+1
								Exit Function
Compila = false
							end if


							arq2.WriteLine("<named name="&chr(34)& elArgs(0) &chr(34)&" helpstring="&chr(34)& elArgs(1) &chr(34)&" required="&chr(34)& elArgs(2) &chr(34)&" type="&chr(34)& "string"&chr(34)&"/>")
							auxDeVerdade = elArgs(0) & "||!-" & elArgs(2)
							args.Add contArgs, split(auxDeVerdade, "||!-")
							nArgs = nArgs + 1
							contArgs = contArgs + 1
						wend

						arq2.WriteLine("</runtime>")

					else
						Message "Erro de sintaxe do comando System Args([name, description, required])" & vbCrlf & "Linha: " & x+1
						Exit Function
Compila = false
					end if
				else
					Message "Erro de sintaxe ou biblioteca inexistente." & vbCrlf & "Linha: " & x+1
					Exit Function
Compila = false
				end if
			else
				Message "Erro de sintaxe ou biblioteca inexistente." & vbCrlf & "Linha: " & x+1
				Exit Function
Compila = false
			end if


			end if

			end if

		elseif NOT linha(x) = "" then
			Message "Erro de sintaxe ou biblioteca inexistente." & vbCrlf & "Linha: " & x+1
			Exit Function
Compila = false
		else
			'Message "aaa"
			end if
		x = x + 1
		
		wend
	
	if UCase(linha5(0)) = "START" and UCase(linha5(1)) = "CODE" then
		arq2.WriteLine"<script language="&chr(34)&"VBScript"&chr(34)&">"
		arq2.WriteLine("On error resume next")
		else
		Message"Erro de sintaxe no inicio do codigo." & vbCrlf & "Linha " & x+1
		Exit Function
Compila = false
	end if
	
	traaab = 0
	linha7 = Split(linha(11))
	if UBound(linha7) = 2 then
	if UCase(linha7(0)) = "VAR" and UCase(linha7(1)) = "STATEMENT" then
		traaab = linha7(2)
		x = 12
		y = 12 + Int(linha7(2))
		while x < y
		
			linhan = Split(linha(x))
			if UBound(linhan) >= 2 then
			
			if UCase(linhan(0)) = "VAR" then
			
			
				if UCase(linhan(1)) = "INTEGER" then
					arq2.WriteLine("Dim " & linhan(2) & "" + vbCrlf + "" & linhan(2) & " = CInt(" & linhan(2) & ")")
					
				elseif UCase(linhan(1)) = "STRING" then 
					arq2.WriteLine("Dim " & linhan(2) & "" + vbCrlf + "" & linhan(2) & " = CStr(" & linhan(2) & ")")
					
				elseif UCase(linhan(1)) = "FLOAT" then
					arq2.WriteLine("Dim " & linhan(2) & "" + vbCrlf + "" & linhan(2) & " = CDbl(" & linhan(2) & ")")

				elseif UCase(linhan(1)) = "DARRAY" then
					arq2.WriteLine("Dim " & linhan(2) & "" + vbCrlf + "Set " & linhan(2) & " = CreateObject(" & chr(34) & "Scripting.Dictionary" & chr(34) & ")")

				elseif UCase(linhan(1)) = "BOOLEAN" then
					arq2.WriteLine("Dim " & linhan(2) & "" + vbCrlf + "" & linhan(2) & " = CBool(" & linhan(2) & ")")

				elseif UCase(linhan(1)) = "DATE" then
					arq2.WriteLine("Dim " & linhan(2) & "" + vbCrlf + "" & linhan(2) & " = CDate(" & linhan(2) & ")")
					
				elseif UCase(linhan(1)) = "SARRAY" then
					if UBound(linhan) = 3 then
						arq2.WriteLine("ReDim " & linhan(2) & "("& linhan(3) &")")
					else
						Message "Tamanho do array nao especificado!" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
					end if
					
				elseif UCase(linhan(1)) = "UNDEFINED" then
					arq2.WriteLine("Dim " & linhan(2))
					
				else 
					Message "Erro na declaracao de variavel." & vbCrlf & "Linha " & x+1
					Exit Function
Compila = false
					end if
					
					
				else
					Message "Erro de Sintaxe no VAR STATEMENT." & vbCrlf & "Linha " & x+1
					Exit Function
Compila = false
					
				end if
			else
				Message"Erro de sintaxe no VAR STATEMENT." & vbCrlf & "Linha " & x+1
				Exit Function
Compila = false
			end if
			x = x + 1
			
			wend
		end if
		else 
			Message"Erro no bloco VAR, estrutura nao reconhecida"
			Exit Function
Compila = false
			end if
			
	linha13 = Split(linha(traaab + 12))
	
	if UBound(linha13) >= 0 then
		Message"Nao sao permitidos caracteres fora dos blocos!" & vbcrlf & "Linha " & traaab+9
		Exit Function
Compila = false
		end if
		
	atlinha = Int(linha7(2)) + 13
	
	linhafunc = Split(linha(atlinha))
	
	atlinha2 = atlinha
	
	' Message atlinha2 & " " & linhafunc(0)
	
	naruto = Split(linha(atlinha2-1))
	
	if UBound(naruto) >= 0 then
		Message"Nao sao permitidos caracteres fora dos blocos!" & vbCrlf & "Linha " & atlinha2
		Exit Function
Compila = false
		end if
		
	linhamain = Split(linha(atlinha2))
	x = atlinha2 + 1
	sedererro = 0
	contaif = 0
	contloop = 0
	deathnumber = totallinhas
	Fores = 0
	if UBound(linhamain) >= 0 then
	if UCase(linhamain(0)) = "MAIN" then
		arq2.WriteLine("Function tudo()")
		arq2.WriteLine("set Comando = WScript.CreateObject(" & chr(34) & "WScript.Shell" & chr(34) & ")	")
		arq2.WriteLine("set FSO = CreateObject("&chr(34)&"Scripting.FileSystemObject"&chr(34)&")")


		while x <= totallinhas

			linha(x) = Replace(linha(x), "|LCASE|", "LCase(")
			linha(x) = Replace(linha(x), "|/CASE|", ")")
			linha(x) = Replace(linha(x), "|UCASE|", "UCase(")
			linha(x) = Replace(linha(x), "|/CASE|", ")")
			linha(x) = Replace(linha(x), "|UFIRST|", "UFirst(")
			linha(x) = Replace(linha(x), "|/CASE|", ")")

			linha(x) = Replace(linha(x), "|LINE|", "vbcrlf")

			if inStr(linha(x), "|THISPATH|") > 0 then
				arq2.WriteLine("thisPathOverHereXd = Comando.CurrentDirectory")
				linha(x) = Replace(linha(x), "|THISPATH|", "thisPathOverHereXd")
			end if

			linhon = Split(linha(x))

			if UBound(linhon)>=0 then


				if NOT UCase(linhon(0)) = "END" and NOT UCase(linhon(0)) = "LOOP" and NOT UCase(linhon(0)) = "IF" and NOT UCase(linhon(0)) = "VAR" and NOT UCase(linhon(0)) = "PRINTVAR" and NOT UCase(linhon(0)) = "PRINT" and NOT UCase(linhon(0)) = "MATH.LIB:EXPO" and NOT UCase(linhon(0)) = "MATH.LIB:SQRT" and NOT UCase(linhon(0)) = "SYSTEMFILES.LIB:WAIT" and NOT UCase(linhon(0)) = "FUNCTION" and NOT UCase(linhon(0)) = "JUMP" and NOT UCase(linhon(0)) = "SYSTEMFILES.LIB:PING" and NOT UCase(linhon(0)) = "SYSTEMFILES.LIB:MACHINE" and NOT UCase(linhon(0)) = "SYSTEMFILES.LIB:OPEN" and NOT UCase(linhon(0)) = "SYSTEMFILES.LIB:MOVE" and NOT UCase(linhon(0)) = "<-" and NOT UCase(linhon(0)) = "READ"  and NOT UCase(linhon(0)) = "//" and NOT UCase(linhon(0)) = "ELSE" and NOT UCase(linhon(0)) = "MATH.LIB:REST" and NOT UCase(linhon(0)) = "BD.LIB:ADDVALUES" and NOT UCase(linhon(0)) = "BD.LIB:GETVALUEROW" and NOT UCase(linhon(0)) = "BD.LIB:USEBD" and NOT UCase(linhon(0)) = "MEDIA.LIB:AUDIO" and NOT UCase(linhon(0)) = "MEDIA.LIB:VIDEO" and NOT UCase(linhon(0)) = "OBJECT" and NOT UCase(linhon(0)) = "WHILE" and NOT UCase(linhon(0)) = "EXIT()" and NOT UCase(linhon(0)) = "BSORT" and NOT UCase(linhon(0)) = "SORT" and NOT UCase(linhon(0)) = "CSVTOVPP" and NOT UCase(linhon(0)) = "VPPTOCSV" and NOT UCase(linhon(0)) = "FOREACH" and NOT UCase(linhon(0)) = "EXITFOR()" and NOT UCase(linhon(0)) = "VARJSON" and NOT UCase(linhon(0)) = "DELETE" then
					if sedererro = 0 then
						Message"Comando Invalido: """ & UCase(linhon(0)) & """." & vbCrlf & "Linha " & x+1
						arq2.WriteLine("<script>")
						Exit Function
Compila = false
						end if
					end if
					
				if x > deathnumber then
					Message"Caracteres invalidos apos finalizacao do bloco Main"
					Arq2.WriteLine("<script>")
					Exit Function
Compila = false 							
					end if
				
					
				
				if UCase(linhon(0)) = "VAR" then
					if UBound(linhon) >= 2 then
						a = 1
						b = ""

						while a <= UBound(linhon)
							podeVar = Split(linha(x), "String(", 2)
							if UBound(podeVar) >= 1 then
								podeVar2 = Split(podeVar(1), ")String", 2)
								if UBound(podeVar2) >= 0 then
									podeVar3 = Split(podeVar2(0), " ")
									if UBound(podeVar3) > 0 then
										Message "Experimente trocar os espacos por _ (underlines) dentro de String(...)String" & vbcrlf & "Linha: " & x+1
										Exit Function
Compila = false
									end if
								end if
							end if

							linhon(a) = Replace(linhon(a), "String(", "" & chr(34))
							linhon(a) = Replace(linhon(a), "_", " ")
							linhon(a) = Replace(linhon(a), "{{", chr(34) & " & ")
							linhon(a) = Replace(linhon(a), "}}", " & " & chr(34))
							linhon(a) = Replace(linhon(a), "\n", chr(34) & " & vbcrlf & " & chr(34) & "")
							linhon(a) = Replace(linhon(a), ")String", "" & chr(34))

							b = b & linhon(a) & " "
							a = a + 1
						wend
						arq2.WriteLine(b)
					else
						Message"Falta de parametros para o metodo VAR" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
					end if

				if UCase(linhon(0)) = "VARJSON" then
					if libjson = 0 then
						Exit Function
Compila = false
					end if

					if UBound(linhon) >= 2 then
						a = 1
						b = ""
						arq2.Write("set ")
						while a <= UBound(linhon)
							podeVar = Split(linha(x), "String(", 2)
							if UBound(podeVar) >= 1 then
								podeVar2 = Split(podeVar(1), ")String", 2)
								if UBound(podeVar2) >= 0 then
									podeVar3 = Split(podeVar2(0), " ")
									if UBound(podeVar3) > 0 then
										Message "Experimente trocar os espacos por _ (underlines) dentro de String(...)String" & vbcrlf & "Linha: " & x+1
										Exit Function
Compila = false
									end if
								end if
							end if

							linhon(a) = Replace(linhon(a), "String(", "" & chr(34))
							linhon(a) = Replace(linhon(a), "_", " ")
							linhon(a) = Replace(linhon(a), "{{", chr(34) & " & ")
							linhon(a) = Replace(linhon(a), "}}", " & " & chr(34))
							linhon(a) = Replace(linhon(a), "\n", chr(34) & " & vbcrlf & " & chr(34) & "")
							linhon(a) = Replace(linhon(a), ")String", "" & chr(34))
							linhon(a) = Lcase(linhon(a))

							b = b & linhon(a) & " "
							a = a + 1
						wend
						arq2.WriteLine(b)
					else
						Message"Falta de parametros para o metodo VAR" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
					end if
		
				if UCase(linhon(0)) = "CSVTOVPP" then
					if libbd = 1 then
						if UBound(linhon) = 1 then
							arq2.WriteLine("CsvToVppBuild(" & chr(34) & linhon(1) & chr(34) & ")")
							arq2.WriteLine("CsvToVppConvert(" & chr(34) & linhon(1) & chr(34) & ")")
						else
							Message "Uso errado da funcao CSVToVpp. Experimente CSVToVpp arquivo" & vbCrlf & "Linha " & x+1
							Exit Function
Compila = false
						end if
					else
						Message "Biblioteca Lib.BD nao importada!" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
					end if
				end if


				if UCase(linhon(0)) = "DELETE" then
					if UBound(linhon) = 1 then
						arq2.WriteLine(linhon(1) & " = Nothing")
					else
						Message "Uso errado do recurso DELETE. Experimente DELETE variavel" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
					end if
				end if



				if UCase(linhon(0)) = "VPPTOCSV" then
					if libbd = 1 then
						if UBound(linhon) = 1 then
							arq2.WriteLine("VppToCsvBuild(" & chr(34) & linhon(1) & chr(34) & ")")
							arq2.WriteLine("VppToCsvConvert(" & chr(34) & linhon(1) & chr(34) & ")")
						else
							Message "Uso errado da funcao VppToCSV. Experimente VppToCSV arquivo" & vbCrlf & "Linha " & x+1
							Exit Function
Compila = false
						end if
					else
						Message "Biblioteca Lib.BD nao importada!" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
					end if
				end if


				
				if UCase(linhon(0)) = "EXITFOR()" then
					if UBound(linhon) = 0 and Fores > 0 then
						arq2.WriteLine("Exit For")
					else
						Message "Uso invalido do comando exitFor()" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
					end if
				end if



				if UCase(linhon(0)) = "BSORT" then
					if libarray = 1 then
						if UBound(linhon) = 2 then
							arq2.WriteLine("Bubble " & linhon(1) & ", " & chr(34) & linhon(2) & chr(34) & " ")
						else
							Message "Uso errado da funcao BSORT. Experimente BSORT arr asc" & vbCrlf & "Linha " & x+1
							Exit Function
Compila = false
						end if
					else
						Message "Biblioteca Lib.Array nao importada!" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
					end if
				end if



				if UCase(linhon(0)) = "SORT" then
					if libarray = 1 then
						if UBound(linhon) = 1 then
							arq2.WriteLine("Bubble " & linhon(1) & ", " & chr(34) & "ASC" & chr(34) & " ")
						else
							Message "Uso errado da funcao SORT. Experimente SORT arr" & vbCrlf & "Linha " & x+1
							Exit Function
Compila = false
						end if
					else
						Message "Biblioteca Lib.Array nao importada!" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
					end if
				end if



				if UCase(linhon(0)) = "OBJECT" then
					if UBound(linhon) = 4 then
						arq2.WriteLine("Set " & linhon(1) & " = New " & linhon(4))
					'elseif UBound(linhon) = 1 then
					'	arq2.WriteLine(linhon(1))
					else
						Message "Erro no instanciamento de classe! Correto: Object a = new classe" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
				end if


					
				if UCase(linhon(0)) = "PRINTVAR" then
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
					arq2.WriteLine("Message " & linhon(1))
					else
						Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
					end if
				if UCase(linhon(0)) = "." then
						MrCatra = "RIP"
					end if
					
				if UCase(linhon(0)) = "READ" then
						if UBound(linhon) = 4 or UBound(linhon) = 3 then
							
							if UBound(linhon) = 4 then
							
								if linhon(1) = "INTEGER" then 
								
									ziri = Replace(linhon(4), "_", " ")
									ziri = Replace(ziri, "{{", chr(34) & " & ")
									ziri = Replace(ziri, "}}", " & " & chr(34))
									
									arq2.WriteLine(linhon(2) & " = Int(Input(""" & ziri & """))")
									ziri = ""
									'arq2.WriteLine(linhon(2) & " = UCase(" & linhon(2) & ")")
								
								elseif linhon(1) = "FLOAT" then
									
									ziri = Replace(linhon(4), "_", " ")
									ziri = Replace(ziri, "{{", chr(34) & " & ")
									ziri = Replace(ziri, "}}", " & " & chr(34))
									
									arq2.WriteLine(linhon(2) & " = cDbl(Input(""" & ziri & """))")
									ziri = ""
									'arq2.WriteLine(linhon(2) & " = UCase(" & linhon(2) & ")")
								
								elseif linhon(1) = "STRING" then
								
									ziri = Replace(linhon(4), "_", " ")
									ziri = Replace(ziri, "{{", chr(34) & " & ")
									ziri = Replace(ziri, "}}", " & " & chr(34))
									
									arq2.WriteLine(linhon(2) & " = Input(""" & ziri & """)")
									ziri = ""
								
								else
									
									Message "Tipo de variavel: " & linhon(2) & " nao encontrado!" & vbCrlf & "Linha: " & x+1
									Exit Function
Compila = false
									
								end if
							
							end if
							
							if UBound(linhon) = 3 then
								ziri = Replace(linhon(3), "_", " ")
								ziri = Replace(ziri, "{{", chr(34) & " & ")
								ziri = Replace(ziri, "}}", " & " & chr(34))
								
								arq2.WriteLine(linhon(1) & " = Input(""" & ziri & """)")
								ziri = ""
							end if
						else
							Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
							Exit Function
Compila = false
						end if
					end if
					
				if UCase(linhon(0)) = "PRINT" then
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
						ziri = Replace(linhon(1), "_", " ")
						ziri = Replace(ziri, "{{", chr(34) & " & ")
						ziri = Replace(ziri, "}}", " & " & chr(34))
						ziri = Replace(ziri, "\n", chr(34) & " & vbcrlf & " & chr(34) & "")

						arq2.WriteLine("Message " & chr(34) & ziri & chr(34))
						ziri = ""
					else
						Message"Sintaxe errada do comando. Experimente trocar os espacos por underlines! (_)" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
					end if
					

				if UCase(linhon(0)) = "WHILE" then
					if UBound(linhon) >= 1 then 
						contez = 1
						arq2.Write("While ")
						while contez <= UBound(linhon)
							arq2.Write(" " & linhon(contez))
							contez = contez + 1
						wend
						contez = 0
						arq2.WriteLine(" ")
					else
						Message "Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
					end if
				end if


				if UCase(linhon(0)) = "MATH.LIB:EXPO" then
				if libmath = 1 then 
					if UBound(linhon) = 4 or UBound(linhon) = 5 then
						arq2.WriteLine(linhon(1) & " = Exponenciacao (" & linhon(3) & "," & linhon(4) & ")")
					else
						Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
						else
						Message "Biblioteca Math nao importada." & vbCrlf & "Linha " & x + 1
						end if
					end if

				if UCase(linhon(0)) = "MATH.LIB:REST" then
				if libmath = 1 then
					if UBound(linhon) = 4 or UBound(linhon) = 5 then
						arq2.WriteLine(linhon(1) & " = Rest (" & linhon(3) & "," & linhon(4) & ")")
					else
						Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
						else
						Message "Biblioteca Math nao importada." & vbCrlf & "Linha " & x + 1
						end if
					end if
				

				if UCase(linhon(0)) = "EXIT()" then
					if UBound(linhon) < 1 then
						arq2.WriteLine("Wscript.quit")
					else
						Message "Uso invalido do comando EXIT." & vbCrlf & "Linha " & x + 1
					end if
				end if



				if UCase(linhon(0)) = "BD.LIB:ADDVALUES" then
				if libbd = 1 then 
					if UBound(linhon) >= 2 then
						happy = 0
						arq2.Write("AddValues (" & chr(34))
						while happy <= UBound(linhon)
							if happy >= 2 then
								arq2.Write(chr(34) & " & " & linhon(happy) & " & " & chr(34) & "")
								if happy = UBound(linhon) then
								else
									arq2.Write("|")
								end if
								else
								arq2.Write(linhon(happy) & " ")
							end if
							happy = happy + 1
						wend
						arq2.Write(chr(34) & ")")
						arq2.Write("" & vbCrlf)
						else
							Message "Sao necessarios mais parametros para inserir na base de dados." & vbCrlf & "Linha " & x + 1
							Exit Function
Compila = false
						end if
					else
						Message "Biblioteca BD nao importada." & vbCrlf & "Linha " & x + 1
						Exit Function
Compila = false
					end if
					end if
					
					
				
				if UCase(linhon(0)) = "MEDIA.LIB:VIDEO" then
					if libmedia = 1 then
						if UBound(linhon) = 1 then
							if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & a & "\Media\" & linhon(1)) then
								openMedia "VIDEO", Comando.CurrentDirectory & "\Projetos\" & a & "\Media\" & linhon(1)
								arq2.WriteLine("if FSO.FileExists(Comando.CurrentDirectory & "&chr(34)&"\Media\" & linhon(1)& chr(34)& ") then")
									arq2.WriteLine("openMedia " & chr(34) & "VIDEO" & chr(34) & ", Comando.CurrentDirectory & "&chr(34)&"\Media\" & linhon(1)& chr(34))
								arq2.WriteLine("end if")
								'arq2.WriteLine("set FSO = CreateObject("&chr(34)&"Scripting.FileSystemObject"&chr(34)&")")
								'arq2.WriteLine("FSO.CopyFile Comando.CurrentDirectory & "&chr(34)&"\Database\" & linhon(1) & ".db"&chr(34)&", Comando.CurrentDirectory & "&chr(34)&"\Projetos\" & a &"\Database\"&chr(34))
							else
								Message "Arquivo " & linhon(1) & " nao encontrado no diretorio de midia do projeto" & vbCrlf & "Linha " & x + 1
								Exit Function
Compila = false
							end if
						else
							Message "Parametros incorretos!" & vbCrlf & "Linha " & x + 1
							Exit Function
Compila = false
						end if
					else
						Message "Biblioteca Media nao importada." & vbCrlf & "Linha " & x + 1
						Exit Function
Compila = false
					end if
				end if
				
				
				if UCase(linhon(0)) = "MEDIA.LIB:AUDIO" then
					if libmedia = 1 then
						if UBound(linhon) = 1 then
							if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & a & "\Media\" & linhon(1)) then
								openMedia "AUDIO", Comando.CurrentDirectory & "\Projetos\" & a & "\Media\" & linhon(1)
								arq2.WriteLine("if FSO.FileExists(Comando.CurrentDirectory & "&chr(34)&"\Media\" & linhon(1)& chr(34)& ") then")
									arq2.WriteLine("openMedia " & chr(34) & "AUDIO" & chr(34) & ", Comando.CurrentDirectory & "&chr(34)&"\Media\" & linhon(1)& chr(34))
								arq2.WriteLine("end if")								'arq2.WriteLine("set FSO = CreateObject("&chr(34)&"Scripting.FileSystemObject"&chr(34)&")")
								'arq2.WriteLine("FSO.CopyFile Comando.CurrentDirectory & "&chr(34)&"\Database\" & linhon(1) & ".db"&chr(34)&", Comando.CurrentDirectory & "&chr(34)&"\Projetos\" & a &"\Database\"&chr(34))
							else
								Message "Arquivo " & linhon(1) & " nao encontrado no diretorio de midia do projeto" & vbCrlf & "Linha " & x + 1
								Exit Function
Compila = false
							end if
						else
							Message "Parametros incorretos!" & vbCrlf & "Linha " & x + 1
							Exit Function
Compila = false
						end if
					else
						Message "Biblioteca Media nao importada." & vbCrlf & "Linha " & x + 1
						Exit Function
Compila = false
					end if
				end if
				
				
				
				if UCase(linhon(0)) = "BD.LIB:USEBD" then
					if libbd = 1 then
						if UBound(linhon) = 1 then
							if FSO.FileExists(Comando.CurrentDirectory & "\Database\" & linhon(1) & ".db") then
								FSO.CopyFile Comando.CurrentDirectory & "\Database\" & linhon(1) & ".db", Comando.CurrentDirectory & "\Projetos\" & a & "\Database\"
								'arq2.WriteLine("set Comando = WScript.CreateObject(" & chr(34) & "WScript.Shell" & chr(34)&")")
								'arq2.WriteLine("set FSO = CreateObject("&chr(34)&"Scripting.FileSystemObject"&chr(34)&")")
								'arq2.WriteLine("FSO.CopyFile Comando.CurrentDirectory & "&chr(34)&"\Database\" & linhon(1) & ".db"&chr(34)&", Comando.CurrentDirectory & "&chr(34)&"\Projetos\" & a &"\Database\"&chr(34))
							else
								Message "Banco de dados " & linhon(1) & " nao encontrado!" & vbCrlf & "Linha " & x + 1
								Exit Function
Compila = false
							end if
						else
							Message "Parametros incorretos!" & vbCrlf & "Linha " & x + 1
							Exit Function
Compila = false
						end if
					else
						Message "Biblioteca BD nao importada." & vbCrlf & "Linha " & x + 1
						Exit Function
Compila = false
					end if
				end if
				
				if UCase(linhon(0)) = "BD.LIB:GETVALUEROW" then
				if libbd = 1 then
					if UBound(linhon) = 4 then
						
						arq2.WriteLine(linhon(4) & " = GetValueRow(" & chr(34) & linhon(1) & chr(34) &", "& linhon(2) & ")")
						
						elseif UBound(linhon) = 5 then
							
						else
							Message "Sao necessarios mais parametros para retornar valores da base de dados." & vbCrlf & "Linha " & x + 1
							Exit Function
Compila = false
						end if
					else
						Message "Biblioteca BD nao importada." & vbCrlf & "Linha " & x + 1
						Exit Function
Compila = false
					end if
					end if
					
				if UCase(linhon(0)) = "SYSTEMFILES.LIB:WAIT" then
				if libsystemfiles = 1 then
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
					arq2.WriteLine("Esperar(" & linhon(1) &")")
						else
						Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
						else 
						Message "Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & x+1
						end if
					end if
					
				if UCase(linhon(0)) = "MATH.LIB:SQRT" then
				if libmath = 1 then
					if UBound(linhon) = 4 or UBound(linhon) = 5 then
					arq2.WriteLine(linhon(1) & " = Raiz (" & linhon(3) & "," & linhon(4) & ")")
					else
						Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false 
						end if
						else
						Message "Biblioteca Math nao importada." & vbCrlf & "Linha " & x + 1
						end if
					end if
					
				if UCase(linhon(0)) = "FUNCTION" then
						if UBound(linhon) >= 1 then
						contez = 1
						if UCase(linhon(1)) = "CALLREQUIREDS(ARGS)" then
							arq2.Writeline("Set args = CreateObject("&chr(34)&"Scripting.Dictionary"&chr(34)&")")
							contadore = 0
							'msgbox nArgs
							while contadore < nArgs
								arq2.WriteLine("straux = " & chr(34) & args(contadore)(0) & "||!-" & args(contadore)(1) & chr(34))
								arq2.WriteLine("args.Add " & contadore & ", split(straux, "&chr(34)&"||!-"&chr(34) & ")")
								contadore = contadore + 1
							wend
						end if

						arq2.Write("Call ")
						while contez <= UBound(linhon)
							arq2.Write(" " & linhon(contez))
							contez = contez + 1
						wend
						contez = 0
						arq2.WriteLine()
						
					else
						Message"Sintaxe errada do comando. Experimento Function funcaoARodar" & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
					end if
				
				if UCase(linhon(0)) = "SYSTEMFILES.LIB:PING" then
				if libsystemfiles = 1 then
					if UBound(linhon) = 2 or UBound(linhon) = 3 then
					arq2.WriteLine("Ping """& linhon(1) &""","&linhon(2)&" ")
					else 
						Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
						else 
						Message "Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & x+1
						end if
					end if
					
				if UCase(linhon(0)) = "SYSTEMFILES.LIB:OPEN" then
				if libsystemfiles = 1 then 
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
					arq2.WriteLine("Abrir ("""& linhon(1) &""")")
					else 
						Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
						else 
						Message "Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & x+1
						end if
					end if
					
				if UCase(linhon(0)) = "SYSTEMFILES.LIB:MACHINE" then
				if libsystemfiles = 1 then
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
					arq2.WriteLine("Maquina("""& linhon(1) &""")")
					else 
						Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
						else 
						Message "Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & x+1
						end if
					end if
					
				if UCase(linhon(0)) = "SYSTEMFILES.LIB:MOVE" then
				if libsystemfiles = 1 then
					if UBound(linhon) = 2 or UBound(linhon) = 3 then
							arq2.WriteLine("Mover"""& linhon(1) &""","""&linhon(2)&"""")
						else 
							Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
							Exit Function
Compila = false
						end if
						else 
						Message "Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & x+1
						end if
					end if
				
				if UCase(linhon(0)) = "IF" then
					if NOT linhon(1) = "=" and linhon(UBound(linhon)) = "->" and NOT linhon(UBound(linhon)-1) = "=" then 
						krai = 1
						pedrinho = 0
						while krai < UBound(linhon)
							if NOT InStr(linhon(krai), "=") = 0 then
								pedrinho = pedrinho + 1
								end if
							krai = krai + 1
						wend
						abc = 1
						arq2.Write("if ")
						while abc < UBound(linhon) 
							linhon(abc) = Replace(linhon(abc), "==", "=")
							linhon(abc) = Replace(linhon(abc), "!=", "<>")
							linhon(abc) = Replace(linhon(abc), "&&", "and")
							linhon(abc) = Replace(linhon(abc), "||", "or")
							arq2.Write(linhon(abc) & " ")
							abc = abc + 1
						wend
							arq2.Write(" then")
							arq2.Writeline("" & vbCrlf)
							contaif = contaif + 1
						else
							Message "Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
							Exit Function
Compila = false
						end if
					end if
				
				if UCase(linhon(0)) = "ELSE" then
					if UBound(linhon) = 0 and contaif > 0then
						arq2.WriteLine("else")
						else
						Message "ELSE nao tem nenhum condicional para referenciar." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
					end if
					
				if UCase(linhon(0)) = "JUMP" then
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
						if linhon(1) = "ERRORS" then 
							sedererro = sedererro + 1
						else 
							Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
							Exit Function
Compila = false
						end if 
						end if
					end if

				if UCase(linhon(0)) = "LOOP" then
					if UBound(linhon) = 2 then
						contifs0 = contaif
						LOOPVAR = LOOPVAR + 1
						arq2.WriteLine("VARLOOP"& LOOPVAR &" = " & linhon(1))
						arq2.WriteLine("REACHVAR"& LOOPVAR &" = " & linhon(2))
						arq2.WriteLine("while VARLOOP" & LOOPVAR & " <= REACHVAR" & LOOPVAR)
						else
						Message"Erro de Sintaxe e/ou estrutura." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
					end if

				if UCase(linhon(0)) = "FOREACH" then
					if UBound(linhon) = 3 then
						'contifs0 = contaif
						'LOOPVAR = LOOPVAR + 1
						arq2.WriteLine("For each " & linhon(1) & " in " & linhon(3))
						Fores = Fores + 1
						else
						Message"Erro de Sintaxe e/ou estrutura." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
					end if


				if UCase(linhon(0)) = "END" then
					if UCase(linhon(1)) = "LOOP" and contifs0 = contaif then
						arq2.WriteLine("VARLOOP"& LOOPVAR & "= VARLOOP"& LOOPVAR &" + 1")
						arq2.WriteLine("wend")
						LOOPVAR = LOOPVAR - 1
						if LOOPVAR < 0 then
							Message"Fechando estrutura inexistente" & vbCrlf & "Linha " & x+1
							arq2.WriteLine("<script>")
							Exit Function
Compila = false
							end if
						elseif UCase(linhon(1)) = "CODE" then
						elseif UCase(linhon(1)) = "WHILE" then
							arq2.WriteLine("wend")
						elseif UCase(linhon(1)) = "FOREACH" then
							if Fores > 0 then
								Fores = Fores - 1
								arq2.WriteLine("Next")
							else
								Message "Fechando Foreach inexistente!" & vbcrlf & "Linha " & x+1
								Exit Function
Compila = false
							end if
						else
						Message"Erro de Sintaxe e/ou estrutura." & vbCrlf & "Linha " & x+1
						Exit Function
Compila = false
						end if
					end if
									
				kaka = 0
				while kaka <= UBound(linhon)
					if linhon(kaka) = "<-" and NOT UCase(linhon(0)) = "//" then
						arq2.WriteLine("end if")
						contaif = contaif - 1
							end if
						if contaif < 0 then
							Message"Fechando estrutura condicional inexistente" & vbCrlf & "Linha " & x+1
							arq2.WriteLine("<script>")
							Exit Function
Compila = false
							end if
					kaka = kaka + 1
					wend
					
					if UCase(linhon(0)) = "END" then
						if UCase(linhon(1)) = "CODE" then
							if UBound(linhon) = 1 then
							deathnumber = x
							arq2.WriteLine("pause("& chr(34) & "Press any key to continue" & chr(34) &")")
							arq2.WriteLine("end function")
							arq2.WriteLine("</script>")
							arq2.WriteLine("<script language = ""JScript"">")
							arq2.WriteLine("try{ ")
							arq2.WriteLine("tudo();")
							arq2.WriteLine("}")
							arq2.WriteLine("catch (e) {")
							arq2.WriteLine("var shell = new ActiveXObject(""Wscript.shell"");")
							arq2.WriteLine("shell.Popup(""Erro de estrutura, divisao por 0 ou erro em valores de variaveis."");")
							arq2.WriteLine("}")
							arq2.WriteLine("</script>")
							arq2.WriteLine("</job>")	
							else
							Message"Erro de Sintaxe e/ou estrutura." & vbCrlf & "Linha " & x+1
							Exit Function
Compila = false
							end if
							if LOOPVAR > 0 or contaif > 0 then
								Message "Estrutura de repeticao ou condicao nao finalizada"
								Exit Function
Compila = false
							end if
							end if
						end if
	
						
				end if
			x = x + 1
			wend	
		else 
			Message"Erro no inicio do bloco MAIN." & vbCrlf & "Linha " & x+1
			Exit Function
Compila = false
			end if
		end if
		

		
	arq2.Close
	'Comando.run"" & a & "_COMPILADO.wsf"
	'Comando.run"cmd /k del C:\ProjetoCompilador\MiniCompilador\" & a & "_COMPILADO.wsf"

	'encoder(Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & "._COMPILADO.wsf")
	'msgbox purinho

	if UCase(codifica) = "SIM" then
		call extractVBS(Comando.CurrentDirectory & "\Projetos\" & purinho & "\", purinho)
		call encodere(Comando.CurrentDirectory & "\Projetos\" & purinho & "\Encoded\", purinho)
		call injectVBE(Comando.CurrentDirectory & "\Projetos\" & purinho & "\", purinho)
	end if

	Compila = true

	end function
	
Function Compexec(a, b, codifica)
	dim aAux
	bAux = a
	aAux = a
	cAux = a
	set Comando = WScript.CreateObject("WScript.Shell")
		if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & aAux & "\" & aAux & ".vpp") then
			Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Projetos\" & aAux & "\" & aAux & ".vpp")	
			if (compila(cAux, codifica)) then
				WScript.Sleep 700
				generateXML(bAux)
				WScript.Sleep 400
				if b = "/CONSOLE" then
					Comando.run "Cscript Projetos\" & aAux & "\" & aAux & "_COMPILADO.wsf"
				elseif b = "/WINDOW" then
					Comando.run "Projetos\" & aAux & "\" & aAux & "_COMPILADO.wsf"
				else
					Comando.run "Projetos\" & aAux & "\" & aAux & "_COMPILADO.wsf"
				end if
			else 
				Message "Compilacao mal-sucedida"
			end if
			
			
		else 
			Message "Arquivo nao encontrado: " & Comando.CurrentDirectory & "\Projetos\" & aAux & "\" & aAux & ".vpp"
			end if
	end function
	
Function limpaLinha(a)
	b = a
	x = 0
	while x < 5
	result1 = InStr(b, "      ")
	result2 = InStr(b, "     ")
	result3 = InStr(b, "    ")
	result4 = InStr(b, "   ")
	result5 = InStr(b, "  ")
	result6 = InStr(b, "	")
	if NOT result1 = 0 or NOT result2 = 0 or NOT result3 = 0 or NOT result4 = 0 or NOT result5 = 0 or NOT result6 = 0 then
		b = Replace(b, "	", " ")
		b = Replace(b, "      ", " ")
		b = Replace(b, "     ", " ")
		b = Replace(b, "    ", " ")
		b = Replace(b, "   ", " ")
		b = Replace(b, "  ", " ")
		end if
		x = x + 1
		wend
	limpaLinha = b
	end function

Function generateXML(a)
	WScript.Sleep 600
	set Comando = WScript.CreateObject("WScript.Shell")
	if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & ".vpp") then
			Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & ".vpp")
			Set arq4 = FSO.CreateTextFile(Comando.CurrentDirectory & "\Projetos\" & a & "\XML\" & a & ".xml", true)
			set arq5 = FSO.OpenTextFile(Comando.CurrentDirectory & "\Config\Config.conf")
			dim vraulios (15)
			zeta = 0
			Do Until arq5.AtEndOfStream
				vraulios(zeta) = arq5.ReadLine
				vraulios(zeta) = UCase(vraulios(zeta))
				vraulios(zeta) = limpaLinha(vraulios(zeta))
				vraulios(zeta) = TRIM(vraulios(zeta))
				zeta = zeta + 1
			Loop
			
			coroi = 1
		arq5.Close
			
	Dim linha (10000)
	y = 0
	LOOPVAR = 0
	Do Until arq.AtEndOfStream
		linha(y) = arq.ReadLine
		linha(y) = UCase(linha(y))
		linha(y) = limpaLinha(linha(y))
		linha(y) = TRIM(linha(y))
		y = y + 1
		Loop
		
	totallinhas = arq.Line-1
	arq.Close
	errolib = 0
	
	
	linha5 = Split(linha(9))
	if UBound(linha5) = 0 then
			Exit Function
		elseif UBound(linha5) = 2 and linha5(0) = "START" and linha5(1) = "CODE" then
			arq4.WriteLine("<XML version=" & chr(34) & "1.0.0" & chr(34) & ">")
			arq4.WriteLine("	<Path> " & Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & ".vpp" & " </Path>")
			while coroi < zeta
				linheun = Split(vraulios(coroi))
				if UBound(linheun) <> 1 then
					Message "Erro no arquivo de configuracao do compilador." & vbCrlf & vbCrlf & "Arquivo: " & Comando.CurrentDirectory & "\Config\Config.conf" & vbCrlf & "Linha: " & coroi + 1
					Exit Function
					end if
				
				if coroi = 1 then
					arq4.Writeline("	<Configuration id=" & chr(34) & "Key" & chr(34) & "> "&  md5(linheun(1)) &" </Configuration>")
					end if
					
				if coroi = 2 then
					arq4.Writeline("	<Configuration id=" & chr(34) & "Creator" & chr(34) & "> "& linheun(1) &" </Configuration>")
					end if
					
				if coroi = 3 then
					arq4.Writeline("	<Configuration id=" & chr(34) & "Token" & chr(34) & "> "& md5(linheun(1)) &" </Configuration>")
					end if
					
				if coroi = 4 then
					arq4.Writeline("	<Configuration id=" & chr(34) & "Launch-Version" & chr(34) & "> "& linheun(1) &" </Configuration>")
					end if
					
				if coroi = 5 then
					arq4.Writeline("	<Configuration id=" & chr(34) & "Version" & chr(34) & "> "& linheun(1) &" </Configuration>")
					end if
					
				if coroi = 6 then
					arq4.Writeline("	<Configuration id=" & chr(34) & "Target-Plataform" & chr(34) & "> "& linheun(1) &" </Configuration>")
					end if
					
				coroi = coroi + 1
			wend
			arq4.Writeline("	<Compilation id=" & chr(34) & "Name" & chr(34) & "> "& a &" </Compilation>")
			arq4.WriteLine("	<Compilation id=" & chr(34) & "Version" & chr(34) & "> 1.9 </Compilation>")
			dataAtual = now()
			dataFormatada = FormatDateTime(dataAtual, 2)
			arq4.WriteLine("	<Compilation id=" & chr(34) & "Date" & chr(34) & "> "& dataFormatada & " </Compilation>")
			arq4.WriteLine("")
			arq4.WriteLine("	<Project id=" & chr(34) & linha5(2) & chr(34) & ">")
		else 
			Exit Function
		end if
	x = 0
	linha6 = Split(linha(10))
	
	if UBound(linha6) >= 0 then
		Exit Function
	end if
	
	libmath = 0
	libsystemfiles = 0
	libbd = 0
	libmedia = 0
	libarray = 0
	while x < 10
	linhasplit = Split(linha(x))
		if linha(x) = "IMPORT LIB.MATH" then
			arq4.WriteLine"		<Import id=" & chr(34) & "Math" & chr(34) & "> </Import>"
			libmath = 1
			
		elseif linha(x) = "IMPORT LIB.SYSTEMFILES" then
			arq4.WriteLine"		<Import id=" & chr(34) & "SystemFiles" & chr(34) & "> </Import>"
			libsystemfiles = 1
			
		elseif linha(x) = "IMPORT LIB.BD" then
			arq4.WriteLine"		<Import id=" & chr(34) & "BD" & chr(34) & "> </Import>"
			libbd = 1

		elseif linha(x) = "IMPORT LIB.ARRAY" then
			arq4.WriteLine"		<Import id=" & chr(34) & "Array" & chr(34) & "> </Import>"
			libarray = 1
			
		elseif linha(x) = "IMPORT LIB.MEDIA" then
			arq4.WriteLine"		<Import id=" & chr(34) & "Media" & chr(34) & "> </Import>"
			libmedia = 1
		
		elseif UBound(linhasplit) >= 0 then
		if linhasplit(0) = "INCLUDE" then
			if UBound(linhasplit) = 1 then
				if FSO.FileExists(Comando.CurrentDirectory & "\Extensoes\" & linhasplit(1) & ".vbs") then
					arq4.WriteLine"		<Import id=" & chr(34) & "Extension" & chr(34) & "> "& linhasplit(1) &" </Import>"
				else
					Exit Function
				end if
			else
				Exit Function
			end if
		end if

		elseif UBound(linhasplit) >= 0 then
		if linhasplit(0) = "CINCLUDE" then
			if UBound(linhasplit) = 1 then
				if FSO.FileExists(Comando.CurrentDirectory & "\Classes\" & linhasplit(1) & ".vbs") then
					arq4.WriteLine"		<Import id=" & chr(34) & "Class" & chr(34) & "> "& linhasplit(1) &" </Import>"
				else
					Exit Function
				end if
			else
				Exit Function
			end if
		end if


		
		elseif NOT linha(x) = "" then

			

		else
			end if
		x = x + 1
		
		wend
		
	if linha5(0) = "START" and linha5(1) = "CODE" then
		else
		Exit Function
	end if
	
	traaab = 0
	linha7 = Split(linha(11))
	if UBound(linha7) = 2 then
	if linha7(0) = "VAR" and linha7(1) = "STATEMENT" then
		arq4.WriteLine("		<Statement id=" & chr(34) & "Var" & chr(34) & ">")
		traaab = linha7(2)
		x = 12
		y = 12 + Int(linha7(2))
		while x < y
		
			linhan = Split(linha(x))
			if UBound(linhan) >= 0 then
			
			if linhan(0) = "VAR" then
			
			
				if linhan(1) = "INTEGER" then
				arq4.WriteLine"			<Var id=" & chr(34) & "Integer" & chr(34) & "> " & linhan(2) & " </Var>"					
				elseif linhan(1) = "STRING" then 
				
				arq4.WriteLine"			<Var id=" & chr(34) & "String" & chr(34) & "> " & linhan(2) & " </Var>"					
					
				elseif linhan(1) = "FLOAT" then
				arq4.WriteLine"			<Var id=" & chr(34) & "Float" & chr(34) & "> " & linhan(2) & " </Var>"					
				
				elseif linhan(1) = "DARRAY" then
				arq4.WriteLine"			<Var id=" & chr(34) & "Dinamic Array" & chr(34) & "> " & linhan(2) & " </Var>"		

				elseif linhan(1) = "BOOLEAN" then
				arq4.WriteLine"			<Var id=" & chr(34) & "Boolean" & chr(34) & "> " & linhan(2) & " </Var>"	

				elseif linhan(1) = "DATE" then				
				arq4.WriteLine"			<Var id=" & chr(34) & "Date" & chr(34) & "> " & linhan(2) & " </Var>"	

				elseif linhan(1) = "SARRAY" then
				arq4.WriteLine"			<Var id=" & chr(34) & "Static Array" & chr(34) & "> " & linhan(2) & " </Var>"	

				elseif linhan(1) = "UNDEFINED" then
				arq4.WriteLine"			<Var id=" & chr(34) & "Undefined" & chr(34) & "> " & linhan(2) & " </Var>"					
					
				else 
					Exit Function
					end if
					
					
				else
					Exit Function
					
				end if
			else
				Exit Function
			end if
			x = x + 1
			
			wend
		end if
		else 
			Exit Function
			end if
			
	linha13 = Split(linha(traaab + 12))
	
	if UBound(linha13) >= 0 then
		Exit Function
		end if
		
	atlinha = Int(linha7(2)) + 13
	linhafunc = Split(linha(atlinha))
		
	atlinha2 = atlinha
	naruto = Split(linha(atlinha2-1))
	
	if UBound(naruto) >= 0 then
		Exit Function
		end if
		
	linhamain = Split(linha(atlinha2))
	x = atlinha2 + 1
	sedererro = 0
	contaif = 0
	contloop = 0
	deathnumber = totallinhas
	if UBound(linhamain) >= 0 then
	if linhamain(0) = "MAIN" then
		arq4.WriteLine("		</Statement>")
		arq4.WriteLine("		<Main>")
		while x <= totallinhas
			linhon = Split(linha(x))
			if UBound(linhon)>=0 then
			
				if NOT UCase(linhon(0)) = "END" and NOT UCase(linhon(0)) = "LOOP" and NOT UCase(linhon(0)) = "IF" and NOT UCase(linhon(0)) = "VAR" and NOT UCase(linhon(0)) = "PRINTVAR" and NOT UCase(linhon(0)) = "PRINT" and NOT UCase(linhon(0)) = "MATH.LIB:EXPO" and NOT UCase(linhon(0)) = "MATH.LIB:SQRT" and NOT UCase(linhon(0)) = "SYSTEMFILES.LIB:WAIT" and NOT UCase(linhon(0)) = "FUNCTION" and NOT UCase(linhon(0)) = "JUMP" and NOT UCase(linhon(0)) = "SYSTEMFILES.LIB:PING" and NOT UCase(linhon(0)) = "SYSTEMFILES.LIB:MACHINE" and NOT UCase(linhon(0)) = "SYSTEMFILES.LIB:OPEN" and NOT UCase(linhon(0)) = "SYSTEMFILES.LIB:MOVE" and NOT UCase(linhon(0)) = "<-" and NOT UCase(linhon(0)) = "READ"  and NOT UCase(linhon(0)) = "//" and NOT UCase(linhon(0)) = "ELSE" and NOT UCase(linhon(0)) = "MATH.LIB:REST" and NOT UCase(linhon(0)) = "BD.LIB:ADDVALUES" and NOT UCase(linhon(0)) = "BD.LIB:GETVALUEROW" and NOT UCase(linhon(0)) = "BD.LIB:USEBD" and NOT UCase(linhon(0)) = "MEDIA.LIB:AUDIO" and NOT UCase(linhon(0)) = "MEDIA.LIB:VIDEO" and NOT UCase(linhon(0)) = "OBJECT" and NOT UCase(linhon(0)) = "WHILE" and NOT UCase(linhon(0)) = "EXIT()" and NOT UCase(linhon(0)) = "BSORT" and NOT UCase(linhon(0)) = "SORT" and NOT UCase(linhon(0)) = "CSVTOVPP" and NOT UCase(linhon(0)) = "VPPTOCSV" and NOT UCase(linhon(0)) = "FOREACH" and NOT UCase(linhon(0)) = "EXITFOR()" and NOT UCase(linhon(0)) = "VARJSON" and NOT UCase(linhon(0)) = "DELETE" then
					if sedererro = 0 then
					Message"erro variaveis"
						Exit Function
						end if
					end if
					
				if x > deathnumber then
				Message "erro deathnumber"
					Exit Function 							
					end if
								
				if UCase(linhon(0)) = "VAR" then
					if UBound(linhon) >= 2 then
						a = 1
						b = ""
						while a <= UBound(linhon)
							b = b & linhon(a) & " "
							a = a + 1
						wend
						arq4.WriteLine("			<Operation> " & b & " </Operation>")
					else
						Message "erro var"
						Exit Function
						end if
					end if
				
				
				if UCase(linhon(0)) = "OBJECT" then
					if UBound(linhon) = 4 then
						arq4.WriteLine("			<Class id=" & chr(34) & "New Instance" & chr(34) & "> " & linhon(1) & " = new " & linhon(4) & " </Class>")
					else
						Exit Function
						end if
				end if

				if UCase(linhon(0)) = "WHILE" then
					if UBound(linhon) >= 1 then 
						contez = 1
						arq4.Write("			<While> ")
						while contez <= UBound(linhon)
							arq4.Write(" " & linhon(contez))
							contez = contez + 1
						wend
						contez = 0
						arq4.WriteLine(" ")
					else
						'Message "Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
					end if
				end if


				if UCase(linhon(0)) = "EXIT()" then
					if UBound(linhon) < 1 then
						arq4.WriteLine("			<Quit id=" & chr(34) & "Application" & chr(34) & "> </Quit>")
					else
						Message "Uso invalido do comando EXIT." & vbCrlf & "Linha " & x + 1
					end if
				end if


				if UCase(linhon(0)) = "PRINTVAR" then
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
						arq4.WriteLine("			<Message id=" & chr(34) & "Var" & chr(34) & "> "& linhon(1) & " </Message>")
					else
					Message"erro printvar"
						Exit Function
						end if
					end if
				if UCase(linhon(0)) = "." then
					end if

				if UCase(linhon(0)) = "CSVTOVPP" then
					if libbd = 1 then
						if UBound(linhon) = 1 then
							arq4.WriteLine("			<Conversion from=" & chr(34) & "CSV" & chr(34) & " to=" & chr(34) & "Vpp" & chr(34) & "> " & linhon(1) & " </Conversion>")
						else
							'Message "Uso errado da funcao CSVToVpp. Experimente CSVToVpp arquivo" & vbCrlf & "Linha " & x+1
							Exit Function
						end if
					else
						'Message "Biblioteca Lib.BD nao importada!" & vbCrlf & "Linha " & x+1
						Exit Function
					end if
				end if

				if UCase(linhon(0)) = "VPPTOCSV" then
					if libbd = 1 then
						if UBound(linhon) = 1 then
							arq4.WriteLine("			<Conversion from=" & chr(34) & "Vpp" & chr(34) & " to=" & chr(34) & "CSV" & chr(34) & "> " & linhon(1) & " </Conversion>")
						else
							'Message "Uso errado da funcao CSVToVpp. Experimente CSVToVpp arquivo" & vbCrlf & "Linha " & x+1
							Exit Function
						end if
					else
						'Message "Biblioteca Lib.BD nao importada!" & vbCrlf & "Linha " & x+1
						Exit Function
					end if
				end if
					
					
				if UCase(linhon(0)) = "BSORT" then
					if libarray = 1 then
						if UBound(linhon) = 2 then
							arq4.WriteLine("			<Sort id=" & chr(34) & "Bubble" & chr(34) & " order=" & chr(34) & linhon(2) & chr(34) & "> "& linhon(1) & " </Sort>")
						else
							'Message "Uso errado da funcao BSORT. Experimente BSORT arr asc" & vbCrlf & "Linha " & x+1
							Exit Function
						end if
					else
						'Message "Biblioteca Lib.Array nao importada!" & vbCrlf & "Linha " & x+1
						Exit Function
					end if
				end if



				if UCase(linhon(0)) = "SORT" then
					if libarray = 1 then
						if UBound(linhon) = 1 then
							arq4.WriteLine("			<Sort id=" & chr(34) & "Bubble" & chr(34) & " order=" & chr(34) & linhon(2) & chr(34) & "> "& linhon(1) & " </Sort>")
						else
							'Message "Uso errado da funcao SORT. Experimente SORT arr" & vbCrlf & "Linha " & x+1
							Exit Function
						end if
					else
						'Message "Biblioteca Lib.Array nao importada!" & vbCrlf & "Linha " & x+1
						Exit Function
					end if
				end if
				
				
				if UCase(linhon(0)) = "READ" then
						if UBound(linhon) = 3 or UBound(linhon) = 4 then
							ziri = Replace(linhon(3), "_", " ")
							arq4.WriteLine("			<Input id=" & chr(34) & linhon(1) & chr(34) & "> " & ziri & " </Input>")
							ziri = ""
						else
						Message "erro read"
							Exit Function
						end if
					end if
					
				if UCase(linhon(0)) = "PRINT" then
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
						ziri = Replace(linhon(1), "_", " ")
						arq4.WriteLine("			<Message id=" & chr(34) & "String" & chr(34) & "> " & ziri & " </Message>")
						ziri = ""
					else
						Message "Erro print"
						Exit Function
					end if
					end if
					
				if UCase(linhon(0)) = "MATH.LIB:EXPO" then
				if libmath = 1 then 
					if UBound(linhon) = 4 or UBound(linhon) = 5 then
						arq4.WriteLine("			<Operation id=" & chr(34) & "LibMath Expo" & chr(34)  & "> " & linhon(1) & " </Operation>")
					else
						Message "erro expo"
						Exit Function
						end if
						else
						end if
					end if

				if UCase(linhon(0)) = "MATH.LIB:REST" then
				if libmath = 1 then
					if UBound(linhon) = 4 or UBound(linhon) = 5 then
						arq4.WriteLine("			<Operation id=" & chr(34) & "LibMath Rest" & chr(34)  & "> " & linhon(1) & " </Operation>")
					else
						Message "erro resto"

						Exit Function
						end if
						else
						end if
					end if
				
				if UCase(linhon(0)) = "BD.LIB:ADDVALUES" then
				if libbd = 1 then 
					if UBound(linhon) >= 2 then
						happy = 0
						fodass = ""
						while happy <= UBound(linhon)
							fodass = fodass & linhon(happy) & " "
							happy = happy + 1
						wend
						arq4.WriteLine("			<Operation id=" & chr(34) & "BD AddValues" & chr(34)  & "> " & fodass & " </Operation>")
						else
						end if
					else
					end if
					end if
				
				if UCase(linhon(0)) = "MEDIA.LIB:VIDEO" then
					if libmedia = 1 then
						if UBound(linhon) = 1 then
							if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & a & "\Media\" & linhon(1)) then
						arq4.WriteLine("			<Operation id=" & chr(34) & "Open Video" & chr(34)  & "> " & linhon(1) & "</Operation>")
							else
								Exit Function
							end if
						else
							Exit Function
						end if
					else
						Exit Function
					end if
				end if
				
				
				if UCase(linhon(0)) = "MEDIA.LIB:AUDIO" then
					if libmedia = 1 then
						if UBound(linhon) = 1 then
							if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & a & "\Media\" & linhon(1)) then
						arq4.WriteLine("			<Operation id=" & chr(34) & "Open Audio" & chr(34)  & "> " & linhon(1) & "</Operation>")
								'arq2.WriteLine("set FSO = CreateObject("&chr(34)&"Scripting.FileSystemObject"&chr(34)&")")
								'arq2.WriteLine("FSO.CopyFile Comando.CurrentDirectory & "&chr(34)&"\Database\" & linhon(1) & ".db"&chr(34)&", Comando.CurrentDirectory & "&chr(34)&"\Projetos\" & a &"\Database\"&chr(34))
							else
							end if
						else
						end if
					else
					end if
				end if
				
				if UCase(linhon(0)) = "BD.LIB:GETVALUEROW" then
				if libbd = 1 then
					if UBound(linhon) = 4 then
						arq4.WriteLine("			<Operation id=" & chr(34) & "BD GetValueRow" & chr(34)  & "> " & linhon(1) & " " & linhon(2) & " </Operation>")
						
						elseif UBound(linhon) = 5 then
							
						else
						end if
					else
					end if
					end if
					
				if UCase(linhon(0)) = "BD.LIB:USEBD" then
					if libbd = 1 then
						if UBound(linhon) = 1 then
							if FSO.FileExists(Comando.CurrentDirectory & "\Database\" & linhon(1) & ".db") then
								arq4.WriteLine("			<Connection id=" & chr(34) & "BD" & chr(34)  & "> " & linhon(1) & " </Connection>")
							else
								Exit Function
							end if
						else
							Exit Function
						end if
					else
						Exit Function
					end if
				end if
				
				if UCase(linhon(0)) = "SYSTEMFILES.LIB:WAIT" then
				if libsystemfiles = 1 then
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
					arq4.WriteLine("			<Operation id=" & chr(34) & "SystemFiles Wait" & chr(34)  & "> " & linhon(1) & " </Operation>")
						else
						Message"Sintaxe errada do comando." & vbCrlf & "Linha " & x+1
						Exit Function
						end if
						else 
						Message "Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & x+1
						end if
					end if
					
				if UCase(linhon(0)) = "MATH.LIB:SQRT" then
				if libmath = 1 then
					if UBound(linhon) = 4 or UBound(linhon) = 5 then
					arq4.WriteLine("			<Operation id=" & chr(34) & "LibMath Sqrt" & chr(34)  & "> " & linhon(1) & " </Operation>")
					else
						Exit Function 
						end if
						else
						end if
					end if
					
				if UCase(linhon(0)) = "FUNCTION" then
					parreira = Split(linhon(1), "(")
					if UBound(linhon) >= 1 then
						arq4.WriteLine("			<Function id=" & chr(34) & "Run" & chr(34)  & "> " & parreira(0) & "() </Function>")
					else
						'Message"Sintaxe errada do comando. Experimento Function funcaoARodar" & vbCrlf & "Linha " & x+1
						Exit Function
						end if
					end if
				
				if UCase(linhon(0)) = "SYSTEMFILES.LIB:PING" then
				if libsystemfiles = 1 then
					if UBound(linhon) = 2 or UBound(linhon) = 3 then
					arq4.WriteLine("			<Operation id=" & chr(34) & "SystemFiles Ping" & chr(34)  & ">")
					arq4.WriteLine("				<Server id=" & chr(34) & "Destination" & chr(34)  & "> " & linhon(1) & " </Server>" )
					arq4.WriteLine("				<Packages id=" & chr(34) & "Number" & chr(34)  & "> " & linhon(2) & " </Packages>" )
					arq4.WriteLine("			</Operation>")
					else 
						Exit Function
						end if
						else 
						end if
					end if
					
				if UCase(linhon(0)) = "SYSTEMFILES.LIB:OPEN" then
				if libsystemfiles = 1 then 
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
					arq4.WriteLine("			<Operation id=" & chr(34) & "LibMath Open" & chr(34)  & "> " & linhon(1) & " </Operation>")
					else 
						Exit Function
						end if
						else 
						end if
					end if
					
				if UCase(linhon(0)) = "SYSTEMFILES.LIB:MACHINE" then
				if libsystemfiles = 1 then
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
					arq4.WriteLine("			<Operation id=" & chr(34) & "LibMath Machine" & chr(34)  & "> " & linhon(1) & " </Operation>")
					else 
						Exit Function
						end if
						else 
						end if
					end if
					
				if UCase(linhon(0)) = "SYSTEMFILES.LIB:MOVE" then
				if libsystemfiles = 1 then
					if UBound(linhon) = 2 or UBound(linhon) = 3 then
						arq4.WriteLine("			<Operation id=" & chr(34) & "LibMath Move" & chr(34)  & "> " & linhon(1))
						arq4.WriteLine("				<File id=" & chr(34) & "Origin" & chr(34)  & "> " & linhon(1) & " </File>")
						arq4.WriteLine("				<File id=" & chr(34) & "Destination" & chr(34)  & "> " & linhon(2) & " </File>")
						arq4.WriteLine("			</Operation>")

						else 
							Exit Function
						end if
						else 
						end if
					end if
				
				if UCase(linhon(0)) = "IF" then
					if NOT linhon(1) = "=" and linhon(UBound(linhon)) = "->" and NOT linhon(UBound(linhon)-1) = "=" then 
						krai = 1
						pedrinho = 0
						while krai < UBound(linhon)
							if NOT InStr(linhon(krai), "=") = 0 then
								pedrinho = pedrinho + 1
								end if
							krai = krai + 1
						wend
						abc = 1
						arq4.WriteLine("			<Conditional id=" & chr(34) & "If" & chr(34)  & ">")
						while abc < UBound(linhon) 
							abc = abc + 1
							wend
							contaif = contaif + 1
						else
							Message "erro no if"
							Exit Function
						end if
					end if
				
				if UCase(linhon(0)) = "ELSE" then
					if UBound(linhon) = 0 and contaif > 0then
						arq4.WriteLine("			</Conditional>")
						arq4.WriteLine("			<Conditional id=" & chr(34) & "Else" & chr(34)  & "> ")
						else
						Message "erro no else"
						Exit Function
						end if
					end if
					
				if UCase(linhon(0)) = "JUMP" then
					if UBound(linhon) = 1 or UBound(linhon) = 2 then
						if linhon(1) = "ERRORS" then 
							arq4.WriteLine("			<Configuration id=" & chr(34) & "Jump errors" & chr(34)  & "> </Configuration>")
							sedererro = sedererro + 1
						else 
							Message "erro no jump"
							Exit Function
						end if 
						end if
					end if

				if UCase(linhon(0)) = "LOOP" then
					if UBound(linhon) = 2 then
						contifs0 = contaif
						LOOPVAR = LOOPVAR + 1
						arq4.WriteLine("			<Repetition id=" & chr(34) & "Loop" & chr(34)  & "> " & linhon(1) & " to " & linhon(2) & " </Repetition>")
						else
						'Message "erro no loop"
						Exit Function
						end if
					end if


				if UCase(linhon(0)) = "FOREACH" then
					if UBound(linhon) = 3 then
						'contifs0 = contaif
						'LOOPVAR = LOOPVAR + 1
						arq4.WriteLine("			<Repetition id=" & chr(34) & "Foreach" & chr(34)  & "> " & linhon(1) & " in " & linhon(3) & " </Repetition>")
						else
						'Message"Erro de Sintaxe e/ou estrutura." & vbCrlf & "Linha " & x+1
						Exit Function
						end if
					end if



				if UCase(linhon(0)) = "END" then
					if linhon(1) = "LOOP" and contifs0 = contaif then
						arq4.WriteLine("			<Repetition id=" & chr(34) & "End" & chr(34)  & "> " & linhon(1) & " </Repetition>")
						LOOPVAR = LOOPVAR - 1
						if LOOPVAR < 0 then
							Exit Function
							end if
						elseif linhon(1) = "CODE" then
						elseif linhon(1) = "WHILE" then
							arq4.WriteLine("			</While>")
						elseif linhon(1) = "FOREACH" then
							arq4.WriteLine("			</Repetition>")
						else
						'Message "erro no end loop"
						Exit Function
						end if
					end if
									
				kaka = 0
				while kaka <= UBound(linhon)
					if linhon(kaka) = "<-" and NOT UCase(linhon(0)) = "//" then
						arq4.WriteLine("			</Conditional>")
						contaif = contaif - 1
							end if
						if contaif < 0 then
							Message "erro end if"
							Exit Function
							end if
					kaka = kaka + 1
					wend
					
					if UCase(linhon(0)) = "END" then
						if linhon(1) = "CODE" then
							if UBound(linhon) = 1 then
							deathnumber = x
							arq4.WriteLine("		</Main>")
							arq4.WriteLine("	</Project>")
							arq4.WriteLine("</XML>")
							else 
							Exit Function
							end if
							if LOOPVAR > 0 or contaif > 0 then
								Exit Function
							end if
							end if
						end if
	
						
				end if
			x = x + 1
			wend	
		else 
			Exit Function
			end if
		end if
		

		
	arq4.Close
	arq.Close
	'Comando.run"" & a & "_COMPILADO.wsf"
	'Comando.run"cmd /k del C:\ProjetoCompilador\MiniCompilador\" & a & "_COMPILADO.wsf"
		else 
			Message "Arquivo nao encontrado: " & Comando.CurrentDirectory & "\Projetos\" & a & "\" & a & ".vpp"
			end if			
	end function


