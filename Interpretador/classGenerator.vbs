'        @PARANAUERJ DEVELOPEMENT WITH UPDATE
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







set FSO = CreateObject("Scripting.FileSystemObject")


Function temNoArray(arr, obj)

	x = 0
	
	indice = -1

    While x <= UBound(arr)
      If arr(x) = obj Then
        indice = x
		x = UBound(arr) + 1
      End If
	  x = x + 1
    Wend
	
  'Err.Clear()
  temNoArray = indice

End Function



Function readClassParams(classe)

	set Comando = WScript.CreateObject("WScript.Shell")
	Set dict = CreateObject("Scripting.Dictionary")
	Set file = FSO.OpenTextFile (Comando.CurrentDirectory & "\Classes\" & classe & ".class.vpp", 1)
	dim linha(10000)

	Do Until file.AtEndOfStream
		linha(y) = file.ReadLine
		' linha(y) = UCase(linha(y))
		linha(y) = limpaLinha(linha(y))
		linha(y) = TRIM(linha(y))
		y = y + 1
	Loop

	totLines = file.Line-1

	file.Close

	cont = 0
	row = 0

	while cont <= totLines
		analisador = Split(linha(cont))

		if UBound(analisador) >= 0 then
			if UCase(analisador(0)) = "SVAR" or UCase(analisador(0)) = "FVAR" then
				if UBound(analisador) = 2 then
					if UCase(analisador(2)) = "PUBLIC" or UCase(analisador(2)) = "PROTECTED" or UCase(analisador(2)) = "PRIVATE" then
						' (nomeVar, Fvar, private)
						dict.Add row, Array(analisador(1), analisador(0), analisador(2))
						row = row + 1
					end if
				end if
			end if
		end if

		cont = cont + 1
	wend

	set readClassParams = dict

End function



Function readClassFuncs(classe)

	set Comando = WScript.CreateObject("WScript.Shell")
	Set dict = CreateObject("Scripting.Dictionary")
	Set file = FSO.OpenTextFile (Comando.CurrentDirectory & "\Classes\Generated\" & classe & ".class.vbs", 1)
	dim linha(10000)

	y = 0
	Do Until file.AtEndOfStream
		linha(y) = file.ReadLine
		y = y + 1
	Loop

	totLines = file.Line-1

	file.Close

	cont = 0
	row = 0
	funcOpen = 0
	fContent = ""

	while cont <= totLines

		analisador = Split(linha(cont))


		if UBound(analisador) >= 0 then

			if UCase(analisador(0)) = "END" then

				if UBound(analisador) = 1 then

					if UCase(analisador(1)) = "FUNCTION" and funcOpen = 1 then
						
						a = dict(row)

						b = Array(a(0), a(1), fContent)

						dict(row) = b

						funcOpen = 0
						fContent = ""
						row = row + 1

					end if

				end if

			end if

			if funcOpen = 1 then

				fContent = fContent & vbcrlf & linha(cont)

			end if
			
			if UBound(analisador) >= 2 then
				if UCase(analisador(1)) = "FUNCTION" then
					
					access = analisador(0)
					nomeFunc = ""

					ieee = 2
					while ieee <= UBound(analisador)
						nomeFunc = nomeFunc & " " & analisador(ieee)
						ieee = ieee + 1
					wend

					' (nomeFunc, public)
					
					dict.Add row, Array(nomeFunc, access)

					funcOpen = 1
						
				end if
			end if


		end if

		cont = cont + 1
	wend

	set readClassFuncs = dict

End function



Function createClass(a)

	set a = fso.CreateTextFile(Comando.CurrentDirectory & "\Classes\" & a & ".class.vpp", true)
		a.WriteLine("Class nome_da_classe")
		a.WriteLine("")
		a.WriteLine("End class")
	a.Close

End Function



Function generateClass(a)
    set Comando = WScript.CreateObject("WScript.Shell")
	Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Classes\" & a & ".class.vpp")	
    'WScript.Sleep 1000
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

    set arq2 = FSO.CreateTextFile(Comando.CurrentDirectory & "\Classes\Generated\" & a & ".class.vbs", true)

    declaracao = 0

    fim = 0

    funcao = 0
    nFunc = ""

    contaif = 0
    uaile = 0
    libmath = 1
	libsystemfiles = 1
	libbd = 1
	libmedia = 1
	libarray = 1
	errors = "ENABLED"
	correctMode = "OFF"
	Fores = 0

	Dim variaveisClass
	Set variaveisClass = CreateObject("Scripting.Dictionary")
	Dim variaveisClassT
	Set variaveisClassT = CreateObject("Scripting.Dictionary")
	contVarC = 0
	Dim variaveisClassS
	Set variaveisClassS = CreateObject("Scripting.Dictionary")
	contVarS = 0
	classname = "?"
	filho = false
	classPai = ""

	' Keywords
	Dim keyWords(41)
	keyWords(0) = "CLASS"
	keyWords(1) = "END"
	keyWords(2) = "WHILE"
	keyWords(3) = "FVAR"
	keyWords(4) = "MATH.LIB:EXPO"
	keyWords(5) = "MATH.LIB:REST"
	keyWords(6) = "FUNCTION"
	keyWords(7) = "VAR"
	keyWords(8) = "PRINTVAR"
	keyWords(9) = "PRINT"
	keyWords(10) = "RETURN"
	keyWords(11) = "IF"
	keyWords(12) = "ELSE"
	keyWords(13) = "CSVTOVPP"
	keyWords(14) = "VPPTOCSV"
	keyWords(15) = "BSORT"
	keyWords(16) = "READ"
	keyWords(17) = "SORT"
	keyWords(18) = "BD.LIB:ADDVALUES"
	keyWords(19) = "MEDIA.LIB:VIDEO"
	keyWords(20) = "MEDIA.LIB:AUDIO"
	keyWords(21) = "BD.LIB:USEBD"
	keyWords(22) = "SYSTEMFILES.LIB:WAIT"
	keyWords(23) = "MATH.LIB:SQRT"
	keyWords(24) = "SYSTEMFILES.LIB:PING"
	keyWords(25) = "SYSTEMFILES.LIB:OPEN"
	keyWords(26) = "SYSTEMFILES.LIB:MACHINE"
	keyWords(27) = "SYSTEMFILES.LIB:MOVE"
	keyWords(28) = "LOOP"
	keyWords(29) = "BD.LIB:GETVALUEROW"
	keyWords(30) = "//"
	keyWords(31) = "<-"
	keyWords(32) = "SETPUBLIC()"
	keyWords(33) = "GETALL()"
	keyWords(34) = "GETPRIVATE()"
	keyWords(35) = "GETPUBLIC()"
	keyWords(36) = "SETALL()"
	keyWords(37) = "SETPRIVATE()"
	keyWords(38) = "SVAR"
	keyWords(39) = "FOREACH"
	keyWords(40) = "EXITFOR()"
	keyWords(41) = "OVERLOAD"
	
	
	'keyWords(38) = "CORRECT()"

    while LOOPVAR <= totallinhas

		linha(LOOPVAR) = Replace(linha(LOOPVAR), "|LINE|", "vbcrlf")
        linhaX = Split(linha(LOOPVAR))

        if UBound(linhaX) >= 0 then 

		if NOT temNoArray(keyWords, UCase(linhaX(0))) = -1 then

            if UCase(linhaX(0)) = "CLASS" and UBound(linhaX) = 1 and declarao = 0 then

                arq2.WriteLine("Class " & linhaX(1))
				classname = linhaX(1)
				
                declaracao = 1

            end if

			if UCase(linhaX(0)) = "CLASS" and UBound(linhaX) = 3 and declarao = 0 then

				if UCase(linhaX(2)) = "EXTENDS" then
					arq2.WriteLine("Class " & linhaX(1))
					classname = linhaX(1)
					classPai = linhaX(3)
					filho = true
					declaracao = 1
					
					' (nomeVar, Fvar, private)
					set paiParam = readClassParams(classPai)
					For Each obj in paiParam.items
						arq2.WriteLine(obj(2) & " " & obj(0))

						variaveisClass.Add contVarC, obj(0)
						variaveisClassT.Add contVarC, obj(2)

						contVarC = contVarC + 1
					Next
					
					' (nomeFunc, private, content)
					set paiFuncs = readClassFuncs(classPai)
					For Each obj in paiFuncs.items
						arq2.WriteLine(obj(1) & " FUNCTION " & obj(0))

						arq2.WriteLine(obj(2))

						arq2.WriteLine("END FUNCTION")
					Next
					
				else
					Message "((CLASSE "&classname&") Operacao invalida! Experimente usar EXTENDS.)" & vbcrlf & "Linha: " & LOOVAR + 1
					Exit function
				end if

            end if

            if declaracao = 0 and fim = 0 then
                Message "(CLASSE "&classname&") Existem elementos fora da classe!" & vbcrlf & "Linha: " & LOOPVAR + 1
                Exit Function
            elseif declaracao = 0 and fim = 1 then
                Message "(CLASSE "&classname&") Nao e possivel fechar uma classe nao declarada!" & vbcrlf & "Linha: " & LOOPVAR + 1
                Exit Function
            end if

			' Correct utilidades
			if correctMode = "ON" then
				if linhaX(0) = "RED" then
					linhaX(0) = Replace(linhaX(0), "RED", "READ")
				end if
			end if
			'End comentário

            if declaracao = 1 then
                if UCase(linhaX(0)) = "END" and UBound(linhaX) = 1 then

                    if UCase(linhaX(1)) = "CLASS" then
                        if funcao = 0 then
							arq2.WriteLine("Private Sub Class_Initialize(  )")
							contadore = 0
							while contadore < contVarS
								arq2.WriteLine("call set" & variaveisClassS(contadore) & "()")
								contadore = contadore + 1
							wend
							'arq2.WriteLine("pause("& chr(34) & "Press any key to continue" & chr(34) &")")
							arq2.WriteLine("End Sub")

                            arq2.WriteLine("End Class")
                            fim = 1
                        else
                             Message "(CLASSE "&classname&") Nao e possivel fechar a classe dentro de uma funcao!" & vbcrlf & "Linha: " & LOOPVAR + 1
                            Exit Function
                        end if
                
                    elseif UCase(linhaX(1)) = "FUNCTION" then
                        if funcao = 1 then
                            arq2.WriteLine("End Function")
                            funcao = 0
                        else
                            Message "(CLASSE "&classname&") Nao e possivel fechar uma funcao que nao foi instanciada!" & vbcrlf & "Linha: " & LOOPVAR + 1
                            Exit Function
                        end if

					elseif UCase(linhaX(1)) = "OVERLOAD" then
                        if funcao = 1 then
                            arq2.WriteLine("End Operator")
                            funcao = 0
                        else
                            Message "(CLASSE "&classname&") Nao e possivel fechar uma funcao que nao foi instanciada!" & vbcrlf & "Linha: " & LOOPVAR + 1
                            Exit Function
                        end if

                    elseif UCase(linhaX(1)) = "WHILE" then
                        if uaile = 1 then
                            arq2.WriteLine("End Function")
                            uaile = uaile - 1
                        else
                            Message "(CLASSE "&classname&") Nao e possivel fechar um while que nao foi instanciado!" & vbcrlf & "Linha: " & LOOPVAR + 1
                            Exit Function
                        end if

                    elseif UCase(linhaX(1)) = "FOREACH" then
						if Fores > 0 then
							arq2.WriteLine("Next")
							Fores = Fores - 1
						else
							Message "(CLASSE "&classname&") Nao e possivel fechar um Foreach que nao foi instanciado!" & vbcrlf & "Linha: " & LOOPVAR + 1
                            Exit Function
						end if
					else
                        Message "(CLASSE "&classname&") Erro no uso de END" & vbcrlf & "Linha: " & LOOPVAR + 1
                        Exit Function
                    end if

                end if

                if UCase(linhaX(0)) = "WHILE" then
					if UBound(linhaX) >= 1 then 
						contez = 1
						arq2.Write("While ")
						while contez <= UBound(linhaX)
							arq2.Write(" " & linhaX(contez))
							contez = contez + 1
						wend
						contez = 0
						arq2.WriteLine(" ")
                        uaile = ualie + 1
					else
						Message "(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if



				if UCase(linhaX(0)) = "CORRECT()" then
					if UBound(linhaX) = 0 then
						errors = "DISABLED"
						correctMode = "ON"
					else
						Message "(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if



				if UCase(linhaX(0)) = "GETALL()" then
					if UBound(linhaX) = 0 then
						cont = 0
						while cont < contVarC
							arq2.WriteLine("Function GET" & variaveisClass(cont) & "()")
								arq2.WriteLine("GET" & variaveisClass(cont) & " = " & variaveisClass(cont))
							arq2.WriteLine("End Function")
							cont = cont + 1
						wend
						cont = 0
					else
						Message "(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if


				if UCase(linhaX(0)) = "SETALL()" then
					if UBound(linhaX) = 0 then
						cont = 0
						while cont < contVarC
							arq2.WriteLine("Function SET" & variaveisClass(cont) & "(param1)")
								arq2.WriteLine(variaveisClass(cont) & " = param1")
							arq2.WriteLine("End Function")
							cont = cont + 1
						wend
						cont = 0
					else
						Message "(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if


				if UCase(linhaX(0)) = "EXITFOR()" then
					if UBound() = 0 and Fores > 0 then
						arq2.WriteLine("Exit For")
					else
						Message "Uso invalido do comando exitFor()" & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if


				if UCase(linhaX(0)) = "FOREACH" then
					if UBound() = 3 then
						'contifs0 = contaif
						'LOOPVAR = LOOPVAR + 1
						arq2.WriteLine("For each " & linhaX(1) & " in " & linhaX(3))
						Fores = Fores + 1
					else
						Message"Erro de Sintaxe e/ou estrutura." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if



				if UCase(linhaX(0)) = "GETPRIVATE()" then
					if UBound(linhaX) = 0 then
						cont = 0
						while cont < contVarC
							if variaveisClassT(cont) = "PRIVATE" then
								arq2.WriteLine("Function GET" & variaveisClass(cont) & "()")
									arq2.WriteLine("GET" & variaveisClass(cont) & " = " & variaveisClass(cont))
								arq2.WriteLine("End Function")
							end if
							cont = cont + 1
						wend
						cont = 0
					else
						Message "(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if


				if UCase(linhaX(0)) = "SETPRIVATE()" then
					if UBound(linhaX) = 0 then
						cont = 0
						while cont < contVarC
							if variaveisClassT(cont) = "PRIVATE" then
								arq2.WriteLine("Function SET" & variaveisClass(cont) & "(param1)")
									arq2.WriteLine(variaveisClass(cont) & " = param1")
								arq2.WriteLine("End Function")
							end if
							cont = cont + 1
						wend
						cont = 0
					else
						Message "(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if


				if UCase(linhaX(0)) = "GETPUBLIC()" then
					if UBound(linhaX) = 0 then
						cont = 0
						while cont < contVarC
							if variaveisClassT(cont) = "PUBLIC" then
								arq2.WriteLine("Function GET" & variaveisClass(cont) & "()")
									arq2.WriteLine("GET" & variaveisClass(cont) & " = " & variaveisClass(cont))
								arq2.WriteLine("End Function")
							end if
							cont = cont + 1
						wend
						cont = 0
					else
						Message "(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if


				if UCase(linhaX(0)) = "SETPUBLIC()" then
					if UBound(linhaX) = 0 then
						cont = 0
						while cont < contVarC
							if variaveisClassT(cont) = "PUBLIC" then
								arq2.WriteLine("Function SET" & variaveisClass(cont) & "(param1)")
									arq2.WriteLine(variaveisClass(cont) & " = param1")
								arq2.WriteLine("End Function")
							end if
							cont = cont + 1
						wend
						cont = 0
					else
						Message "(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if



                if UCase(linhaX(0)) = "FVAR" then
                    if UBound(linhaX) = 2 then
                        arq2.WriteLine(linhaX(2) & " " & linhaX(1))
						variaveisClass.Add contVarC, linhaX(1)
						variaveisClassT.Add contVarC, linhaX(2)
						contVarC = contVarC + 1
                    else
                        Message "(CLASSE "&classname&") Erro em declaracao de variavel! O correto e: FVar name type" & vbcrlf & "Linha: " & LOOPVAR + 1
                        Exit Function
                    end if
                end if



				if UCase(linhaX(0)) = "SVAR" then
                    if UBound(linhaX) = 3 then
						if linhaX(1) = "SARRAY" then
							arq2.WriteLine(linhaX(3) & " " & linhaX(2))
						end if
						if linhaX(1) = "DARRAY" then
							arq2.WriteLine(linhaX(3) & " " & linhaX(2))
							arq2.WriteLine("Function set" & linhaX(2) & "()")
							arq2.WriteLine("Set " & linhaX(2) & " = CreateObject(" & chr(34) & "Scripting.Dictionary" & chr(34) & ")")
							arq2.WriteLine("End Function")
							variaveisClassS.Add contVarS, linhaX(2)
							contVarS = contVarS + 1
						end if
                        
						'variaveisClass.Add contVarC, linhaX(1)
						'variaveisClassT.Add contVarC, linhaX(2)
						'contVarC = contVarC + 1
                    else
                        Message "(CLASSE "&classname&") Erro em declaracao de variavel! O correto e: SVar tipoDeSpecialVar name type" & vbcrlf & "Linha: " & LOOPVAR + 1
                        Exit Function
                    end if
                end if



                if UCase(linhaX(0)) = "MATH.LIB:EXPO" then
				if libmath = 1 then 
					if UBound(linhaX) = 4 or UBound(linhaX) = 5 then
						arq2.WriteLine(linhaX(1) & " = Exponenciacao (" & linhaX(3) & "," & linhaX(4) & ")")
					else
						Message"(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
						end if
						else
						Message "(CLASSE "&classname&") Biblioteca Math nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						end if
					end if

				if UCase(linhaX(0)) = "MATH.LIB:REST" then
				if libmath = 1 then
					if UBound(linhaX) = 4 or UBound(linhaX) = 5 then
						arq2.WriteLine(linhaX(1) & " = Rest (" & linhaX(3) & "," & linhaX(4) & ")")
					else
						Message"(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
						end if
						else
						Message "(CLASSE "&classname&") Biblioteca Math nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						end if
					end if

                

                if UCase(linhaX(0)) = "FUNCTION" then
                    if funcao = 0 and UBound(linhaX) >= 2 then
						If NOT UCase(linhaX(1)) = "PUBLIC" and NOT UCase(linhaX(1)) = "PRIVATE" then
							Message "(CLASSE "&classname&") Erro na declaracao da funcao! O tipo da funcao deve ser apenas Public ou Private!" & vbcrlf & "Linha: " & LOOPVAR + 1
							Exit Function
						end if
                        arq2.Write(linhaX(1) & " Function")
                        cont = 2
                        while cont <= UBound(linhaX)
                            arq2.Write(" " & linhaX(cont))
                            cont = cont + 1
                        wend
                        cont = 0

                        arq2.WriteLine(" ")
                        FnameAux = Split(linhaX(2), "(")
                        nFunc = FnameAux(0)
                        funcao = 1
                    else
                        Message "(CLASSE "&classname&") Erro na declaracao da funcao! O correto e: function type fName(parameters)" & vbcrlf & "Linha: " & LOOPVAR + 1
                        Exit Function
                    end if
                end if


				if UCase(linhaX(0)) = "OVERLOAD" then
                    if funcao = 0 and UBound(linhaX) >= 4 then
						cont = 3
                        while cont <= UBound(linhaX)
                            textoImp = textoImp & " " & linhaX(cont)
							cont = cont + 1
                        wend
                        cont = 0

						textoImp = replace(textoImp, "(", "")
						textoImp = replace(textoImp, ")", "")
						textoImp = TRIM(textoImp)

						textoCmp = split(textoImp, " ")

						if(UBound(textoCmp) > 1) then
							Message "Erro no overload!" & vbcrlf & "Linha: " & LOOPVAR
							Exit Function
						end if

						tpToCmp = textoCmp(0)
						nmToCmp = textoCmp(1)

                        arq2.Write(linhaX(1) & " Public Shared Operator " & linhaX(2) & " (ByVal Value As " & classname & ", ByVal " & nmToCmp & " As " & tpToCmp & ") As " & classname & "")
                        arq2.WriteLine(" ")
                        FnameAux = Split(linhaX(2), "(")
                        nFunc = FnameAux(0)
                        funcao = 1
                    else
                        Message "(CLASSE "&classname&") Erro na declaracao do operador! O correto e: overload operator + (parameters)" & vbcrlf & "Linha: " & LOOPVAR + 1
                        Exit Function
                    end if
                end if


                if UCase(linhaX(0)) = "VAR" then
					if UBound(linhaX) >= 2 then
						a = 1
						b = ""

						while a <= UBound(linhaX)

							podeVar = Split(linha(LOOPVAR), "String(", 2)
							if UBound(podeVar) >= 1 then
								podeVar2 = Split(podeVar(1), ")String", 2)
								if UBound(podeVar2) >= 0 then
									podeVar3 = Split(podeVar2(0), " ")
									if UBound(podeVar3) > 0 then
										Message "(CLASSE "&classname&") Experimente trocar os espacos por _ (underlines) dentro de String(...)String" & vbcrlf & "Linha: " & LOOPVAR + 1
										Exit Function
									end if
								end if
							end if

							linhaX(a) = Replace(linhaX(a), "String(", "" & chr(34))
							linhaX(a) = Replace(linhaX(a), "_", " ")
							linhaX(a) = Replace(linhaX(a), "{{", chr(34) & " & ")
							linhaX(a) = Replace(linhaX(a), "}}", " & " & chr(34))
							linhaX(a) = Replace(linhaX(a), "\n", chr(34) & " & vbcrlf & " & chr(34) & "")
							linhaX(a) = Replace(linhaX(a), ")String", "" & chr(34))

							b = b & linhaX(a) & " "
							a = a + 1
						wend
						arq2.WriteLine(b)
					else
						Message "(CLASSE "&classname&") Falta de parametros para o metodo VAR" & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
                    end if
                end if



                if UCase(linhaX(0)) = "PRINTVAR" then
					if UBound(linhaX) = 1 then
					    arq2.WriteLine("Message " & linhaX(1))
					else
						Message "(CLASSE "&classname&") Sintaxe errada do comando PrintVar" & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
                    end if
                end if



            if UCase(linhaX(0)) = "PRINT" then
                if UBound(linhaX) = 1 or UBound(linhaX) = 2 then
                    ziri = Replace(linhaX(1), "_", " ")
                    ziri = Replace(ziri, "{{", chr(34) & " & ")
                    ziri = Replace(ziri, "}}", " & " & chr(34))
                    ziri = Replace(ziri, "\n", chr(34) & " & vbcrlf & " & chr(34) & "")

                    arq2.WriteLine("Message " & chr(34) & ziri & chr(34))
                    ziri = ""
                else
                    Message"(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
                    Exit Function
                end if
            end if



            if UCase(linhaX(0)) = "RETURN" then
                if UBound(linhaX) = 1 and funcao = 1 then
                    arq2.WriteLine(nFunc & " = " & linhaX(1))
                else
                    Message"(CLASSE "&classname&") Erro no uso do Return!" & vbCrlf & "Linha " & LOOPVAR + 1
                    Exit Function
                end if
            end if



            if UCase(linhaX(0)) = "IF" then
					if NOT linhaX(1) = "=" and linhaX(UBound(linhaX)) = "->" and NOT linhaX(UBound(linhaX)-1) = "=" then 
						krai = 1
						pedrinho = 0
						while krai < UBound(linhaX)
							if NOT InStr(linhaX(krai), "=") = 0 then
								pedrinho = pedrinho + 1
								end if
							krai = krai + 1
						wend
						abc = 1
						arq2.Write("if ")
						while abc < UBound(linhaX) 
                            linhaX(abc) = Replace(linhaX(abc), "==", "=")
							linhaX(abc) = Replace(linhaX(abc), "!=", "<>")
							linhaX(abc) = Replace(linhaX(abc), "&&", "and")
							linhaX(abc) = Replace(linhaX(abc), "||", "or")
							arq2.Write(linhaX(abc) & " ")
							abc = abc + 1
							wend
							arq2.Write(" then")
							arq2.Writeline("" & vbCrlf)
							contaif = contaif + 1
						else
							Message "(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
							Exit Function
						end if
					end if
				
				if UCase(linhaX(0)) = "ELSE" then
					if UBound(linhaX) = 0 and contaif > 0then
						arq2.WriteLine("else")
						else
						Message "(CLASSE "&classname&") ELSE nao tem nenhum condicional para referenciar." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
						end if
					end if

            if UCase(linhaX(0)) = "CSVTOVPP" then
					if libbd = 1 then
						if UBound(linhaX) = 1 then
							arq2.WriteLine("CsvToVppBuild(" & chr(34) & linhaX(1) & chr(34) & ")")
							arq2.WriteLine("CsvToVppConvert(" & chr(34) & linhaX(1) & chr(34) & ")")
						else
							Message "(CLASSE "&classname&") Uso errado da funcao CSVToVpp. Experimente CSVToVpp arquivo" & vbCrlf & "Linha " & LOOPVAR + 1
							Exit Function
						end if
					else
						Message "(CLASSE "&classname&") Biblioteca Lib.BD nao importada!" & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if



				if UCase(linhaX(0)) = "VPPTOCSV" then
					if libbd = 1 then
						if UBound(linhaX) = 1 then
							arq2.WriteLine("VppToCsvBuild(" & chr(34) & linhaX(1) & chr(34) & ")")
							arq2.WriteLine("VppToCsvConvert(" & chr(34) & linhaX(1) & chr(34) & ")")
						else
							Message "(CLASSE "&classname&") Uso errado da funcao VppToCSV. Experimente VppToCSV arquivo" & vbCrlf & "Linha " & LOOPVAR + 1
							Exit Function
						end if
					else
						Message "(CLASSE "&classname&") Biblioteca Lib.BD nao importada!" & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if



				if UCase(linhaX(0)) = "BSORT" then
					if libarray = 1 then
						if UBound(linhaX) = 2 then
							arq2.WriteLine("Bubble " & linhaX(1) & ", " & chr(34) & linhaX(2) & chr(34) & " ")
						else
							Message "(CLASSE "&classname&") Uso errado da funcao BSORT. Experimente BSORT arr asc" & vbCrlf & "Linha " & LOOPVAR + 1
							Exit Function
						end if
					else
						Message "(CLASSE "&classname&") Biblioteca Lib.Array nao importada!" & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if



                if UCase(linhaX(0)) = "READ" then
						if UBound(linhaX) = 4 or UBound(linhaX) = 3 then
							
							if UBound(linhaX) = 4 then
							
								if linhaX(1) = "INTEGER" then 
								
									ziri = Replace(linhaX(4), "_", " ")
									ziri = Replace(ziri, "{{", chr(34) & " & ")
									ziri = Replace(ziri, "}}", " & " & chr(34))
									
									arq2.WriteLine(linhaX(2) & " = Int(Input(""" & ziri & """))")
									ziri = ""
									'arq2.WriteLine(linhaX(2) & " = UCase(" & linhaX(2) & ")")
								
								elseif linhaX(1) = "FLOAT" then
									
									ziri = Replace(linhaX(4), "_", " ")
									ziri = Replace(ziri, "{{", chr(34) & " & ")
									ziri = Replace(ziri, "}}", " & " & chr(34))
									
									arq2.WriteLine(linhaX(2) & " = cDbl(Input(""" & ziri & """))")
									ziri = ""
									'arq2.WriteLine(linhaX(2) & " = UCase(" & linhaX(2) & ")")
								
								elseif linhaX(1) = "STRING" then
								
									ziri = Replace(linhaX(4), "_", " ")
									ziri = Replace(ziri, "{{", chr(34) & " & ")
									ziri = Replace(ziri, "}}", " & " & chr(34))
									
									arq2.WriteLine(linhaX(2) & " = Input(""" & ziri & """)")
									ziri = ""
								
								else
									
									Message "(CLASSE "&classname&") Tipo de variavel: " & linhaX(2) & " nao encontrado!" & vbCrlf & "Linha: " & LOOPVAR + 1
									Exit Function
									
								end if
							
							end if
							
							if UBound(linhaX) = 3 then
								ziri = Replace(linhaX(3), "_", " ")
								ziri = Replace(ziri, "{{", chr(34) & " & ")
								ziri = Replace(ziri, "}}", " & " & chr(34))
								
								arq2.WriteLine(linhaX(1) & " = Input(""" & ziri & """)")
								ziri = ""
							end if
						else
							Message"(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
							Exit Function
						end if
					end if




				if UCase(linhaX(0)) = "SORT" then
					if libarray = 1 then
						if UBound(linhaX) = 1 then
							arq2.WriteLine("Bubble " & linhaX(1) & ", " & chr(34) & "ASC" & chr(34) & " ")
						else
							Message "(CLASSE "&classname&") Uso errado da funcao SORT. Experimente SORT arr" & vbCrlf & "Linha " & LOOPVAR + 1
							Exit Function
						end if
					else
						Message "(CLASSE "&classname&") Biblioteca Lib.Array nao importada!" & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if



            kaka = 0
            while kaka <= UBound(linhaX)
                if linhaX(kaka) = "<-" and NOT UCase(linhaX(0)) = "//" then
                    arq2.WriteLine("end if")
                    contaif = contaif - 1
                        end if
                    if contaif < 0 then
                        Message"(CLASSE "&classname&") Fechando estrutura condicional inexistente" & vbCrlf & "Linha " & LOOPVAR + 1
                        arq2.WriteLine("<script>")
                        Exit Function
                        end if
                kaka = kaka + 1
                wend



            if UCase(linhaX(0)) = "BD.LIB:ADDVALUES" then
                if UBound(linhaX) >= 2 then
                    happy = 0
                    arq2.Write("AddValues (" & chr(34))
                    while happy <= UBound(linhaX)
						linhaX(happy) = Replace(linhaX(happy), "{{", chr(34) & " & ")
						linhaX(happy) = Replace(linhaX(happy), "}}", " & " & chr(34))
                        if happy >= 2 then
                            arq2.Write(chr(34) & " & " & linhaX(happy) & " & " & chr(34) & "")
                            if happy = UBound(linhaX) then
                            else
									arq2.Write("|")
                            end if
                            else
                            arq2.Write(linhaX(happy) & " ")
                        end if
                        happy = happy + 1
                    wend
                    arq2.Write(chr(34) & ")")
                    arq2.Write("" & vbCrlf)
                    else
                        Message "(CLASSE "&classname&") Sao necessarios mais parametros para inserir na base de dados." & vbCrlf & "Linha " & LOOPVAR + 1
                        Exit Function
                    end if
            end if


            if UCase(linhaX(0)) = "MEDIA.LIB:VIDEO" then
					if libmedia = 1 then
						if UBound(linhaX) = 1 then
							if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & a & "\Media\" & linhaX(1)) then
								openMedia "VIDEO", Comando.CurrentDirectory & "\Projetos\" & a & "\Media\" & linhaX(1)
								arq2.WriteLine("if FSO.FileExists(Comando.CurrentDirectory & "&chr(34)&"\Media\" & linhaX(1)& chr(34)& ") then")
									arq2.WriteLine("openMedia " & chr(34) & "VIDEO" & chr(34) & ", Comando.CurrentDirectory & "&chr(34)&"\Media\" & linhaX(1)& chr(34))
								arq2.WriteLine("end if")
								'arq2.WriteLine("set FSO = CreateObject("&chr(34)&"Scripting.FileSystemObject"&chr(34)&")")
								'arq2.WriteLine("FSO.CopyFile Comando.CurrentDirectory & "&chr(34)&"\Database\" & linhaX(1) & ".db"&chr(34)&", Comando.CurrentDirectory & "&chr(34)&"\Projetos\" & a &"\Database\"&chr(34))
							else
								Message "(CLASSE "&classname&") Arquivo " & linhaX(1) & " nao encontrado no diretorio de midia do projeto" & vbCrlf & "Linha " & LOOPVAR + 1
								Exit Function
							end if
						else
							Message "(CLASSE "&classname&") Parametros incorretos!" & vbCrlf & "Linha " & LOOPVAR + 1
							Exit Function
						end if
					else
						Message "(CLASSE "&classname&") Biblioteca Media nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if
				

            if UCase(linhaX(0)) = "MEDIA.LIB:AUDIO" then
					if libmedia = 1 then
						if UBound(linhaX) = 1 then
							if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & a & "\Media\" & linhaX(1)) then
								openMedia "AUDIO", Comando.CurrentDirectory & "\Projetos\" & a & "\Media\" & linhaX(1)
								arq2.WriteLine("if FSO.FileExists(Comando.CurrentDirectory & "&chr(34)&"\Media\" & linhaX(1)& chr(34)& ") then")
									arq2.WriteLine("openMedia " & chr(34) & "AUDIO" & chr(34) & ", Comando.CurrentDirectory & "&chr(34)&"\Media\" & linhaX(1)& chr(34))
								arq2.WriteLine("end if")								'arq2.WriteLine("set FSO = CreateObject("&chr(34)&"Scripting.FileSystemObject"&chr(34)&")")
								'arq2.WriteLine("FSO.CopyFile Comando.CurrentDirectory & "&chr(34)&"\Database\" & linhaX(1) & ".db"&chr(34)&", Comando.CurrentDirectory & "&chr(34)&"\Projetos\" & a &"\Database\"&chr(34))
							else
								Message "(CLASSE "&classname&") Arquivo " & linhaX(1) & " nao encontrado no diretorio de midia do projeto" & vbCrlf & "Linha " & LOOPVAR + 1
								Exit Function
							end if
						else
							Message "(CLASSE "&classname&") Parametros incorretos!" & vbCrlf & "Linha " & LOOPVAR + 1
							Exit Function
						end if
					else
						Message "(CLASSE "&classname&") Biblioteca Media nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if


            if UCase(linhaX(0)) = "BD.LIB:USEBD" then
					if libbd = 1 then
						if UBound(linhaX) = 1 then

                            arq2.WriteLine("FSO.CopyFile Comando.CurrentDirectory & " & chr(34) & "\Database\" & chr(34) & " & " & linhaX(1) & " & " & chr(34) & ".db" & chr(34) & ", Comando.CurrentDirectory & " & chr(34) & "\Projetos\"  & chr(34) & " & a & " & chr(34) & "\Database\" & chr(34))
                            'FSO.CopyFile Comando.CurrentDirectory & "\Database\" & linhaX(1) & ".db", Comando.CurrentDirectory & "\Projetos\" & a & "\Database\"


                            'arq2.WriteLine("set Comando = WScript.CreateObject(" & chr(34) & "WScript.Shell" & chr(34)&")")
                            'arq2.WriteLine("set FSO = CreateObject("&chr(34)&"Scripting.FileSystemObject"&chr(34)&")")
                            'arq2.WriteLine("FSO.CopyFile Comando.CurrentDirectory & "&chr(34)&"\Database\" & linhaX(1) & ".db"&chr(34)&", Comando.CurrentDirectory & "&chr(34)&"\Projetos\" & a &"\Database\"&chr(34))
                        
						else
							Message "(CLASSE "&classname&") Parametros incorretos!" & vbCrlf & "Linha " & LOOPVAR + 1
							Exit Function
						end if
					else
						Message "(CLASSE "&classname&") Biblioteca BD nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
					end if
				end if

            

            if UCase(linhaX(0)) = "SYSTEMFILES.LIB:WAIT" then
				if libsystemfiles = 1 then
					if UBound(linhaX) = 1 or UBound(linhaX) = 2 then
					arq2.WriteLine("Esperar(" & linhaX(1) &")")
						else
						Message"(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
						end if
						else 
						Message "(CLASSE "&classname&") Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						end if
					end if
					
				if UCase(linhaX(0)) = "MATH.LIB:SQRT" then
				if libmath = 1 then
					if UBound(linhaX) = 4 or UBound(linhaX) = 5 then
					arq2.WriteLine(linhaX(1) & " = Raiz (" & linhaX(3) & "," & linhaX(4) & ")")
					else
						Message"(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function 
						end if
						else
						Message "(CLASSE "&classname&") Biblioteca Math nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						end if
					end if


                if UCase(linhaX(0)) = "SYSTEMFILES.LIB:PING" then
				if libsystemfiles = 1 then
					if UBound(linhaX) = 2 or UBound(linhaX) = 3 then
					arq2.WriteLine("Ping """& linhaX(1) &""","&linhaX(2)&" ")
					else 
						Message"(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
						end if
						else 
						Message "(CLASSE "&classname&") Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						end if
					end if
					
				if UCase(linhaX(0)) = "SYSTEMFILES.LIB:OPEN" then
				if libsystemfiles = 1 then 
					if UBound(linhaX) = 1 or UBound(linhaX) = 2 then
					arq2.WriteLine("Abrir ("""& linhaX(1) &""")")
					else 
						Message"(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
						end if
						else 
						Message "(CLASSE "&classname&") Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						end if
					end if
					
				if UCase(linhaX(0)) = "SYSTEMFILES.LIB:MACHINE" then
				if libsystemfiles = 1 then
					if UBound(linhaX) = 1 or UBound(linhaX) = 2 then
					arq2.WriteLine("Maquina("""& linhaX(1) &""")")
					else 
						Message"(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
						end if
						else 
						Message "(CLASSE "&classname&") Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						end if
					end if
					
				if UCase(linhaX(0)) = "SYSTEMFILES.LIB:MOVE" then
				if libsystemfiles = 1 then
					if UBound(linhaX) = 2 or UBound(linhaX) = 3 then
							arq2.WriteLine("Mover"""& linhaX(1) &""","""&linhaX(2)&"""")
						else 
							Message"(CLASSE "&classname&") Sintaxe errada do comando." & vbCrlf & "Linha " & LOOPVAR + 1
							Exit Function
						end if
						else 
						Message "(CLASSE "&classname&") Biblioteca SystemFiles nao importada." & vbCrlf & "Linha " & LOOPVAR + 1
						end if
					end if

                
                if UCase(linhaX(0)) = "LOOP" then
					if UBound(linhaX) = 2 then
						contifs0 = contaif
						LOOPVAR = LOOPVAR + 1
						arq2.WriteLine("VARLOOP"& LOOPVAR &" = " & linhaX(1))
						arq2.WriteLine("REACHVAR"& LOOPVAR &" = " & linhaX(2))
						arq2.WriteLine("while VARLOOP" & LOOPVAR & " <= REACHVAR" & LOOPVAR)
						else
						Message"(CLASSE "&classname&") Erro de Sintaxe e/ou estrutura." & vbCrlf & "Linha " & LOOPVAR + 1
						Exit Function
						end if
					end if


                

            if UCase(linhaX(0)) = "BD.LIB:GETVALUEROW" then
                    if UBound(linhaX) = 4 then
                        
                        arq2.WriteLine(linhaX(4) & " = GetValueRow(" & linhaX(1) & ", "& linhaX(2) & ")")
                        
                        elseif UBound(linhaX) = 5 then
                            
                        else
                            Message "(CLASSE "&classname&") Sao necessarios mais parametros para retornar valores da base de dados." & vbCrlf & "Linha " & LOOPVAR + 1
                            Exit Function
                        end if
            end if




            end if

			else

				Message "(CLASSE "&classname&") Comando invalido: " & UCase(linhaX(0)) & vbcrlf & "Linha: " & LOOPVAR + 1
				Exit Function

			end if
        end if


    LOOPVAR = LOOPVAR + 1

    wend

    if declaracao = 0 then
        Message "(CLASSE "&classname&") O nome da classe nao foi declarado!"
    end if

End Function