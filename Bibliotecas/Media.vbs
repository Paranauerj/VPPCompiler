'        @PARANAUERJ DEVELOPEMENT
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







set FSO = CreateObject("Scripting.FileSystemObject")
set WShell = WScript.CreateObject("WScript.Shell")
set Comando = WScript.CreateObject("WScript.Shell")


function openMedia(tipo, caminhoParam)
	
	if(tipo = "AUDIO" or tipo = "VIDEO") then
		set comando = CreateObject("WScript.shell")
		Set TypeLib = CreateObject("Scriptlet.TypeLib")
		uniqid = TypeLib.Guid
		uniqid = Left(uniqid , Len(uniqid )-2)
		caminho = uniqid & ".vbs"
		caminho = Replace(caminho,"{","")
		caminho = Replace(caminho,"}","")
		caminho = Replace(caminho,"-","")
		Dim fso, MyFile
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set MyFile = fso.CreateTextFile(caminho, True)
	end if
	
	
	
	if(tipo = "AUDIO") then
		
		MyFile.WriteLine("Set fso = CreateObject(" & chr(34) & "Scripting.FileSystemObject" & chr(34) & ")")
		MyFile.WriteLine("Set oPlayer = CreateObject(" & chr(34) & "WMPlayer.OCX" & chr(34) & ")")
		MyFile.WriteLine("oPlayer.URL = " & chr(34) & caminhoParam & chr(34) & "")
		MyFile.WriteLine("oPlayer.controls.play")
		MyFile.WriteLine("While oPlayer.playState <> 1 ' 1 = Stopped")
		MyFile.WriteLine("WScript.Sleep 100")
		MyFile.WriteLine("Wend")
		MyFile.WriteLine("oPlayer.close")
		MyFile.WriteLine("fso.DeleteFile " & chr(34) & caminho & chr(34) & "")


		MyFile.Close

		comando.run ""&caminho


	elseif (tipo = "VIDEO") then
		MyFile.WriteLine("Set fso = CreateObject(" & chr(34) & "Scripting.FileSystemObject" & chr(34) & ")")
		MyFile.WriteLine("Set oPlayer = CreateObject(" & chr(34) & "WMPlayer.OCX" & chr(34) & ")")
		MyFile.WriteLine("oPlayer.openPlayer(" & chr(34) & caminhoParam & chr(34) & ")")
		MyFile.WriteLine("fso.DeleteFile " & chr(34) & caminho & chr(34) & "")
		
		
		MyFile.Close
		
		
		comando.run ""&caminho
	
	else


		msgbox "Formato " & tipo & " Invalido! Tente: play audio FAIXA.mp3 ou play video VIDEO.mp4"
	
	end if
	
	

end function

function displayImg(largura, altura, url, titulo, marginTop, marginLeft)

	Set objExplorer = CreateObject("InternetExplorer.Application")

	With objExplorer
		.Navigate "about:blank"
		.ToolBar = 0
		.StatusBar = 0
		.Left = marginLeft
		.Top = marginTop
		'.Width = 525
		.Width = largura + 30
		'.Height = 555
		.Height = altura + 55
		.Visible = 1
		.Document.Title = "Titulo"
		.Document.Body.InnerHTML = "<img src='" & url & "' height='" & altura & "px' width='" & largura & "px'>"
	End With


end function

function displayVideo(largura, altura, url, titulo, marginTop, marginLeft)

	Set objExplorer = CreateObject("InternetExplorer.Application")

	With objExplorer
		.Navigate "about:blank"
		.ToolBar = 0
		.StatusBar = 0
		.Left = marginLeft
		.Top = marginTop
		'.Width = 525
		.Width = largura + 30
		'.Height = 555
		.Height = altura + 55
		.Visible = 1
		.Document.Title = "Titulo"
		.Document.Body.InnerHTML = "<video src='" & url & "' height='" & altura & "px' width='" & largura & "px' controls> "
	End With


end function


function displayIframe(largura, altura, url, titulo, marginTop, marginLeft)

	Set objExplorer = CreateObject("InternetExplorer.Application")

	With objExplorer
		.Navigate "about:blank"
		.ToolBar = 0
		.StatusBar = 0
		.Left = marginLeft
		.Top = marginTop
		'.Width = 525
		.Width = largura + 35
		'.Height = 555
		.Height = altura + 60
		.Visible = 1
		.Document.Title = titulo
		.Document.Body.InnerHTML = "<iframe src='" & url & "' height='" & altura & "px' width='" & largura & "px'> </iframe>"
	End With


end function

Function getResolution()
	Dim objDictionary
	Set objDictionary = CreateObject("Scripting.Dictionary")
	objDictionary.CompareMode = vbTextCompare
	Set oIE = CreateObject("InternetExplorer.Application")
	With oIE
		.Navigate("about:blank")
		Do Until .readyState = 4: wscript.sleep 100: Loop
				objDictionary.Add "WIDTH", .document.ParentWindow.screen.width
				objDictionary.Add "HEIGHT", .document.ParentWindow.screen.height
	End With
 
	oIE.Quit
	SET getResolution = objDictionary
End Function