'        @PARANAUERJ DEVELOPEMENT
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







Function rodarTerminal()
	a = true
	while a = true
	comando = InputBox("Terminal: " & vbCrLf & "Comando:%>" & vbCrLf & vbCrLf & vbCrLf & "Para saber mais comandos, digite Help" & vbCrLf & vbCrLf, "ParanaShell")
	if comando = "" then
		msgbox"Valores nulos enviados!" + vbCrLf + "Para fechar o terminal, insira o comando 'sair'"
	else
		interpretaComando(comando)	
		end if
	wend
	end function

