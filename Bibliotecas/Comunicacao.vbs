'        @PARANAUERJ DEVELOPEMENT UPDATED 2
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







Function Diga(a)
	msgbox "Saida: " + vbCrLf + "" & a
	end function
	
	
Function Informacao()
	msgbox "Os comandos a seguir funcionam apenas no terminal. Para saber mais sobre os comandos da linguagem, abra Info/Comandos.txt"
	msgbox "Comandos da Biblioteca Math: " + vbCrLf + "Somar var1 var2" + vbCrLf + "Subtrair var1 var2" +vbCrLf + "Multiplicar var1 var2" +vbCrLf + "Dividir var1 var2" +vbCrLf + "Exponenciacao base expoente" +vbCrLf + "Raiz numero num.daraiz" +vbCrLf + "Delta var1 var2 var3" +vbCrLf + "Bhaskara var1 var2 var3" +vbCrLf + "DistanciaPontos xa ya xb yb"
	msgbox "Comandos da Biblioteca SystemFiles: " +vbCrLf + "Mover origem destino" +vbCrLf + "Copiar origem destino" +vbCrLf + "Deletar arquivo" +vbCrLf + "Sair" + vbCrLf +"Ping servidor numero_de_pacotes" + vbCrLf +"Traceroute servidor" + vbCrLf +"Reinciar: Reinicia a aplicacao" +vbCrLf + "MatarTask task" + vbCrLf +"Maquina comando(desligar, reiniciar, hibernar, logoff) "+ vbCrLf+ "CriarProjeto nome_do_projeto" + vbCrLf + "Compila nome_do_projeto" + vbCrLf + "DIGA TEXTO" + vbCrLf + "ABRIR ARQUIVO" + vbCrLf + "PESQUISAR O_QUE_DESEJA" + vbCrLf + "EXEC PROJETO" + vbCrLf + "EDITAR PROJETO" + vbCrLf + "EXPORTARPROJETO NOMEDOPROJETO PASTADESTINO" + vbCrLf + "IMPORTARPROJETO NOMEDOPROJETO PASTAORIGEM" + vbCrLf + "REDE" + vbCrLf + "CRIARPROJETO NOME" + vbCrLf + "COMPILA NOME" + vbCrLf + "ADD EXTENSAO" + vbCrLf + "COMPEXEC NOME" + vbCrLf + "GENERATEXML NOME_DO_PROJETO" + vbCrLf + "PROJETOS" + vbCrLf + "EXTENSOES" + vbCrLf + "ATUALIZAR" + vbCrLf + "PRINT PALAVRA" + vbCrLf + "BuildClass Classe (na pasta Classes da raiz)"
	msgbox "Comandos da Biblioteca Media: " + vbCrLf + "MEDIA TIPODEMEDIA(AUDIO OU VIDEO) ARQUIVO"
	msgbox "Comandos da Biblioteca BD: " + vbCrLf + "ADDBD NOME" + vbCrLf + "ADDROW DATABASE NOME" + vbCrLf + "ADDVALUES DATABASE CAMPOS" + vbCrLf + "GETROWS DATABASE" + vbCrLf + "GETVALUES DATABASE" + vbCrLf + "GETVALUEROW DATABASE NUMERO_DO_ROW"	 + vbCrLf + "CSVTOVPP arquivo(na pasta Database da raiz)" + vbCrLf + "VPPTOCSV arquivo(na pasta Database da raiz)"
	msgbox "Outros comandos: " + vbCrLf + "Reiniciar" + vbCrLf + "Sair"
	end function
