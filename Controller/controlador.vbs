'        @PARANAUERJ DEVELOPEMENT
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







set FSO = CreateObject("Scripting.FileSystemObject")

Function interpretaComando(a)
	set Comando = WScript.CreateObject("WScript.Shell")
	c = Split(UCase(a))
	if c(0) = "SAIR" or c(0) = "EXIT" then
		if UBound(c) = 0 then
			Sair()
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: SAIR"
			end if
		
		
	elseif c(0) = "MOVER" or c(0) = "MOVE" then
		if UBound(c) = 2 then
			msgbox Mover(c(1),c(2))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: MOVER N1 N2"
			end if
		
	
	elseif c(0) = "COPIAR" or c(0) = "COPY" then
		if UBound(c) = 2 then
			msgbox Copiar(c(1),c(2))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: COPIAR ORIGEM DESTINO"
			end if
		
		
	elseif c(0) = "DELETAR" or c(0) = "DELETE" then
		if UBound(c) = 1 then
			msgbox Deletar(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: DELETAR CAMINHO"
			end if
		
		
	elseif c(0) = "SOMAR" or c(0) = "SUM" then
	if UBound(c) = 2 then
		msgbox Somar(c(1),c(2))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: SOMAR N1 N2"
			end if
		
		
	elseif c(0) = "SUBTRAIR" or c(0) = "SUBTRACT" then
		if UBound(c) = 2 then 
			msgbox Subtrair(c(1),c(2))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: SUBTRAIR N1 N2"
			end if
	
	
	elseif c(0) = "MULTIPLICAR" or c(0) = "MULTIPLY" then
		if UBound(c) = 2 then
			msgbox Multiplicar(c(1),c(2))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: MULTIPLICAR N1 N2"
			end if
	
	
	elseif c(0) = "DIVIDIR" or c(0) = "DIVIDE" then
		if UBound(c) = 2 then 
			msgbox Dividir(c(1),c(2))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: DIVIDIR N1 N2"
			end if
	
	
	elseif c(0) = "EXPONENCIACAO" or c(0) = "EXPONENTIATION" then
		if UBound(c) = 2 then 
			msgbox Exponenciacao(c(1),c(2))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: EXPONENCIACAO N1 N2"
			end if
	
	
	elseif c(0) = "RAIZ" or c(0) = "ROOT" then
		if UBound(c) = 2 then 
			msgbox Raiz(c(1),c(2))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: RAIZ N1 N2"
			end if
	
	
	elseif c(0) = "DELTA" then
		if UBound(c) = 3 then
			msgbox Delta(c(1),c(2),c(3))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: DELTA A B C"
			end if
	
	
	elseif c(0) = "BHASKARA" then
		if UBound(c) = 3 then
			msgbox ("X': " & Bhaskara(c(1),c(2),c(3))(0) & "" + vbCrLf + "X'': " & Bhaskara(c(1),c(2),c(3))(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: BHASKARA N1 N2 N3"
			end if
		
		
	elseif c(0) = "DISTANCIAPONTOS" or c(0) = "POINTSDISTANCE" then
		if UBound(c) = 4 then
			msgbox DistanciaPontos(c(1),c(2),c(3),c(4))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: DISTANCIAPONTOS N1 N2 N3 N4"
			end if
		
		
	elseif c(0) = "PING" then 	
		if UBound(c) = 2 then
			Ping c(1),c(2)
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: PING IP PACOTES"
			end if
		
		
	elseif c(0) = "TRACEROUTE" then 	
		if UBound(c) = 1 then
			Traceroute(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: TRACEROUTE IP"
			end if
		
		
	elseif c(0) = "REINICIAR" or c(0) = "RESTART" then 
		if UBound(c) = 0 then 
			Reiniciar()
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: REINICIAR"
			end if
		
		
	elseif c(0) = "MATARTASK" or c(0) = "TASKKILL" then
		if UBound(c) = 1 then
			matarTask(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: MATARTASK TASK"
			end if
		
		
	elseif c(0) = "MAQUINA" or c(0) = "MACHINE" then
		if UBound(c) = 1 then
			Maquina(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: MAQUINA COMANDO"
			end if
		
		
	elseif c(0) = "PROGRAMAR" or c(0) = "PROGRAM" then
		if UBound(c) = 0 then
			Programar()
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: PROGRAMAR (NAO DISPONIVEL)"
			end if
		
		
	elseif c(0) = "DIGA" or c(0) = "SAY" then
		tamanho = UBound(c)
		parametro = ""
		x = 1
		while x <= tamanho
			parametro = parametro & " " & c(x)
			x = x + 1
			wend
		Diga(parametro)
		
		
	elseif c(0) = "INFO" or c(0) = "HELP" or c(0) = "AJUDA" then
		if UBound(c) = 0 then
			Informacao()
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: AJUDA/INFO/HELP"
			end if
		
		
	elseif c(0) = "ABRIR" or c(0) = "OPEN" then
		if UBound(c) = 1 then
			Abrir(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: ABRIR ARQUIVO"
			end if
		
	elseif c(0) = "PESQUISAR" or c(0) = "SEARCH" then
		if UBound(c) = 1 then
			Pesquisar(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: PESQUISAR CHAVE_DE_PESQUISA"
			end if
		
		
	elseif c(0) = "EXEC" then
	if UBound(c) = 1 then
			call AbrirProjeto(c(1), "nada", "")
		elseif UBound(c) = 2 then
			if c(2) = "/CONSOLE" or c(2) = "/WINDOW" then
				call AbrirProjeto(c(1), c(2), "")
			else
				msgbox "Sequencia errada! O parametro opcional (terceiro) deve ser /Console ou /Window!" & vbcrlf & "Correto: EXEC NOME [/Console or /Window]"
			end if
		elseif UBound(c) >= 3 then
			if c(2) = "/CONSOLE" or c(2) = "/WINDOW" then
				contaA = 0
				tP = ""
				while contaA <= UBound(c)
					tP = tP & " " & c(contaA)
					contaA = contaA + 1
				wend
				call AbrirProjeto(c(1), c(2), tP)
			else
				msgbox "Sequencia errada! O parametro opcional (terceiro) deve ser /Console ou /Window!" & vbcrlf & "Correto: EXEC NOME [/Console or /Window]"
			end if
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: EXEC NOME [/Console or /Window] [ARGS]"
			end if
	

			


			
	elseif c(0) = "EDITAR" or c(0) = "EDIT" then
		if UBound(c) = 1 then
			EditarProjeto(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: EDITAR PROJETO"
			end if
			
			
	elseif c(0) = "EXPORTARPROJETO" or c(0) = "EXPORTPROJECT" then
		if UBound(c) = 2 then
			ExportarProjeto c(1), c(2)
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: EXPORTARPROJETO NOMEDOPROJETO PASTADESTINO"
			end if
		
	
	elseif c(0) = "IMPORTARPROJETO" or c(0) = "IMPORTPROJECT" then
		if UBound(c) = 2 then
			ImportarProjeto c(1), c(2)
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: IMPORTARPROJETO NOMEDOPROJETO PASTAORIGEM"
			end if
	
	
	elseif c(0) = "MEDIA" then
		if UBound(c) = 2 then
			openMedia c(1), c(2)
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: MEDIA TIPODEMEDIA(AUDIO OU VIDEO) ARQUIVO"
			end if
			
		
	elseif c(0) = "REDE" or c(0) = "NETWORK" then 
		if UBound(c) = 0 then
			Rede()
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: REDE"
			end if
		
		
	elseif c(0) = "CRIARPROJETO" or c(0) = "CREATEPROJECT" then
		if UBound(c) = 1 then
			CriarArquivo (c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: CRIARPROJETO NOME"
			end if
		
		
	elseif c(0) = "COMPILA" or c(0) = "COMPILE" then
		if UBound(c) = 1 then
			if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & c(1) & "\" & c(1) & ".vpp") then
				call Compila(c(1), "NAO")
			else
				msgbox "Arquivo nao encontrado: " & Comando.CurrentDirectory & "\" & c(1) & ".vpp"
				end if
		elseif UBound(c) = 2 then
			if FSO.FileExists(Comando.CurrentDirectory & "\Projetos\" & c(1) & "\" & c(1) & ".vpp") then
				if c(2) = "/ENCODED" then
					call Compila(c(1), "SIM")
				else
					msgbox "O parametro opcional para esta funcao deve ser /encoded ou nenhum outro!"
				end if
			else
				msgbox "Arquivo nao encontrado: " & Comando.CurrentDirectory & "\" & c(1) & ".vpp"
				end if
		else
			msgbox "Sequencia errada! " & vbcrlf & "Correto: COMPILA NOME"
			end if
	

	elseif c(0) = "CREATECLASS" then
		if UBound(c) = 1 then
			if NOT FSO.FileExists(Comando.CurrentDirectory & "\Classes\" & c(1) & ".class.vpp") then
				createClass(c(1))
			else 
				msgbox "Ja existe uma classe com este nome: " & Comando.CurrentDirectory & "\Classes\" & c(1) & ".class.vpp"
			end if
		else
			msgbox "Sequencia errada! " & vbcrlf & "Correto: BUILD CLASSNAME"
		end if


	elseif c(0) = "BUILDCLASS" then
		if UBound(c) = 1 then
			if FSO.FileExists(Comando.CurrentDirectory & "\Classes\" & c(1) & ".class.vpp") then
				generateClass(c(1))
			else 
				msgbox "Arquivo nao encontrado: " & Comando.CurrentDirectory & "\Classes\" & c(1) & ".class.vpp"
			end if
		else
			msgbox "Sequencia errada! " & vbcrlf & "Correto: BUILD CLASSNAME"
		end if


	elseif c(0) = "ADD" then
		if UBound(c) = 1 then
			addExtension(c(1))
		else
			msgbox "Sequencia errada!" & vbCrLf & "Correto: ADD EXTENSAO"
		end if
	
	elseif c(0) = "COMPEXEC" then
		if UBound(c) = 1 then
			call Compexec(c(1), "nada", "NAO")
		elseif UBound(c) = 2 then
			if c(2) = "/CONSOLE" or c(2) = "/WINDOW" then
				call Compexec(c(1), c(2), "NAO")
			else
				msgbox "Sequencia errada! O parametro opcional (terceiro) deve ser /Console ou /Window!" & vbcrlf & "Correto: COMPEXEC NOME [/Console or /Window]"
			end if
		elseif UBound(c) = 3 then
			if c(2) = "/CONSOLE" or c(2) = "/WINDOW" then
				if c(3) = "/ENCODED" then
					call Compexec(c(1), c(2), "SIM")
				else
					msgbox "O terceiro parametro deve ser /encoded ou nenhum outro!"
				end if
			else
				msgbox "Sequencia errada! O parametro opcional (terceiro) deve ser /Console ou /Window!" & vbcrlf & "Correto: COMPEXEC NOME [/Console or /Window]"
			end if
		else
			msgbox "Sequencia errada! " & vbcrlf & "Correto: COMPEXEC NOME [/Console or /Window]"
			end if
	
	
	elseif c(0) = "ADDBD" then
		if UBound(c) = 1 then
			CriaBD(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: ADDBD NOME"
			end if
			
			
	elseif c(0) = "CSVTOVPP" then
		if UBound(c) = 1 then
			CsvToVppBuild(c(1))
			CsvToVppConvert(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: CSVTOVPP NOME"
			end if


	elseif c(0) = "VPPTOCSV" then
		if UBound(c) = 1 then
			VppToCsvBuild(c(1))
			VppToCsvConvert(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: CSVTOVPP NOME"
			end if

			
			
	elseif c(0) = "ADDROW" then
		if UBound(c) = 2 then
			AddRow c(1), c(2)
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: ADDROW DATABASE NOME"
			end if
			

	'elseif c(0) = "AIROW" then
	'	if UBound(c) = 2 then
	'		AiRow c(1), c(2)
	'	else 
	'		msgbox "Sequencia errada! " & vbcrlf & "Correto: AIROW DATABASE NOME"
	'		end if
			
			
	elseif c(0) = "ADDVALUES" then
		if UBound(c) > 1 then
			AddValues(UCASE(a))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: ADDVALUES DATABASE CAMPOS"
			end if
			
			
	elseif c(0) = "GETROWS" then
		if UBound(c) = 1 then
			pimba = GetRows(c(1))
			msgbox pimba
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: GETROWS DATABASE"
			end if
			
			
	elseif c(0) = "GETVALUES" then
		if UBound(c) = 1 then
			pimba = GetValues(c(1))
			msgbox pimba
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: GETVALUES DATABASE"
			end if
			
			
	elseif c(0) = "GETVALUEROW" then
		if UBound(c) = 2 then
			pimba = GetValueRow(c(1), c(2))
			lll = 0
			stringRows = ""
			while lll <= UBound(pimba)
				stringRows = stringRows & " " & pimba(lll)
				lll = lll + 1
			wend
			msgbox "Valores do row: " & vbcrlf & TRIM(stringRows)
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: GETVALUEROW DATABASE NUMERO_DO_ROW"
			end if
			
			
	elseif c(0) = "GENERATEXML" then
		if UBound(c) = 1 then
			GenerateXML(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: GENERATEXML NOME_DO_PROJETO"
			end if
			
			
	elseif c(0) = "PROJETOS" or c(0) = "PROJECTS" then
		if UBound(c) = 0 then
			Projetos()
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: PROJETOS"
			end if


	elseif c(0) = "EXTENSOES" or c(0) = "EXTENSIONS" then
		if UBound(c) = 0 then
			Extensoes()
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: EXTENSOES"
			end if

	
	elseif c(0) = "UPDATE" or c(0) = "ATUALIZAR" then
		if UBound(c) = 0 then
			tryUpdate()
		end if
			
	elseif c(0) = "PRINT" then
		if UBound(c) = 1 then 
			msgbox(c(1))
		else 
			msgbox "Sequencia errada! " & vbcrlf & "Correto: PRINT PALAVRA"
			end if
	
	
	
	else
		exibeMensagem 1
 	end if
	
	
	if   Err.Number <> 0 then
		exibeMensagem 2
		Err.Clear
	else
		end If
	end function