'        @PARANAUERJ DEVELOPEMENT
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







set Comando = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set arq = FSO.CreateTextFile(Comando.CurrentDirectory & "\Config\Config.xml", true)
vbsversion = fso.GetFileVersion("c:\windows\system32\wscript.exe")

arq.WriteLine("<XML version=" & chr(34) & "1.0.0" & chr(34) &">")
arq.WriteLine("	<VBS id=" & chr(34) & "Version" & chr(34) &"> "& vbsversion & " </VBS>")
arq.WriteLine("	<Instalation id=" & chr(34) & "Path" & chr(34) &"> " & Comando.CurrentDirectory & "\ </Instalation>")
arq.WriteLine("	<Product id=" & chr(34) & "Version" & chr(34) &"> 1.9 </Product>")
arq.WriteLine("	<Product id=" & chr(34) & "Version Compilation Date" & chr(34) &"> 15/05/2019 </Product>")
arq.WriteLine("	<Product id=" & chr(34) & "Initial Compilation" & chr(34) &"> 12/07/2018 </Product>")

pasta = Comando.CurrentDirectory & "\Extensoes\"
For each arquivo in FSO.GetFolder(pasta).Files
	projetos = Replace(arquivo, Comando.CurrentDirectory & "\Extensoes\", " ")
	arq.WriteLine("	<Extension id=" & chr(34) & "Include" & chr(34) &"> " & projetos & " </Extension>")
Next

set arq5 = FSO.OpenTextFile(Comando.CurrentDirectory & "\Config\Config.conf")
dim vraulios (15)
zeta = 0
Do Until arq5.AtEndOfStream
	vraulios(zeta) = arq5.ReadLine
	vraulios(zeta) = UCase(vraulios(zeta))
	vraulios(zeta) = TRIM(vraulios(zeta))
	zeta = zeta + 1
	Loop
	
coroi = 1
arq5.Close
			while coroi < zeta
				linheun = Split(vraulios(coroi))
				if UBound(linheun) <> 1 then
					msgbox "Erro no arquivo de configuracao do compilador." & vbCrlf & vbCrlf & "Arquivo: " & Comando.CurrentDirectory & "\Config\Config.conf" & vbCrlf & "Linha: " & coroi + 1
					WScript.Quit
					end if
				
				if coroi = 1 then
					arq.Writeline("	<Configuration id=" & chr(34) & "Key" & chr(34) & "> "& md5(linheun(1)) &" </Configuration>")
					end if
					
				if coroi = 2 then
					arq.Writeline("	<Configuration id=" & chr(34) & "Creator" & chr(34) & "> "& linheun(1) &" </Configuration>")
					end if
					
				if coroi = 3 then
					arq.Writeline("	<Configuration id=" & chr(34) & "Token" & chr(34) & "> "& md5(linheun(1)) &" </Configuration>")
					end if
					
				if coroi = 4 then
					arq.Writeline("	<Configuration id=" & chr(34) & "Launch-Version" & chr(34) & "> "& linheun(1) &" </Configuration>")
					end if
					
				if coroi = 5 then
					arq.Writeline("	<Configuration id=" & chr(34) & "Version" & chr(34) & "> "& linheun(1) &" </Configuration>")
					end if
					
				if coroi = 6 then
					arq.Writeline("	<Configuration id=" & chr(34) & "Target-Plataform" & chr(34) & "> "& linheun(1) &" </Configuration>")
					end if
					
				coroi = coroi + 1
			wend

arq.WriteLine("</XML>")
arq.Close