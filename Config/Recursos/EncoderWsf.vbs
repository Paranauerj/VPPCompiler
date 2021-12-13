Function encodere(project, a)

    project = project & a & ".vbs"

	set oFSO = CreateObject("Scripting.FileSystemObject")
	
	if oFSO.FileExists(project) then
	dim oEncoder, oFilesToEncode, file, sDest
	dim sFileOut, oFile, oEncFile, oFSO, i
	dim oStream, sSourceFile
	 
	set oEncoder = CreateObject("Scripting.Encoder")
	
	file = project
	
	set oFile = oFSO.GetFile(file)
	Set oStream = oFile.OpenAsTextStream(1)
	sSourceFile=oStream.ReadAll
	oStream.Close
	sDest = oEncoder.EncodeScriptFile(".vbs",sSourceFile,0,"")
	sFileOut = Left(file, Len(file) - 3) & "vbe"
	Set oEncFile = oFSO.CreateTextFile(sFileOut)
	oEncFile.Write sDest
	oEncFile.Close
	else
		msgbox "Nao achei!"
	end if
	
	oFSO.deleteFile project
	
End Function

Function extractVBS(project, a)
	
	set fso = CreateObject("Scripting.FileSystemObject")
	'Set dict = CreateObject("Scripting.Dictionary")
	'msgbox "path: " & project & a & "_COMPILADO.wsf"
	Set file = fso.OpenTextFile(project & a & "_COMPILADO.wsf", 1)
	
	'project = replace(project, ".wsf", ".vbs")

	

	Set arq = fso.CreateTextFile(project & "Encoded\" & a & ".vbs", True)
	
	
	

	'row = 0
	abriu = 0
	pode = 0
	
	Do Until file.AtEndOfStream
	
	  line = file.Readline
	  line = TRIM(line)
	  
	if UCase(line) = "<SCRIPT LANGUAGE=" & chr(34) & "VBSCRIPT" & chr(34) & ">" then
		pode = 1
	end if
	
	
	if UCase(line) = "</SCRIPT>" then
		exit do
	elseif abriu = 1 then
		arq.WriteLine(line)
		'dict.Add row, line
		'row = row + 1
	end if
	
	if pode = 1 then
		pode = 0
		abriu = 1
	end if
	
	Loop

	file.Close
	arq.Close

end function

Function injectVBE(project, a)

    projectEn = a & ".vbe"
    project = project & a & "_COMPILADO.wsf"
	
	
	set fso = CreateObject("Scripting.FileSystemObject")
	Set dict = CreateObject("Scripting.Dictionary")
	Set file = fso.OpenTextFile(project, 1)
	
	row = 0
	abriu = 0
	pode = 1
	
	Do Until file.AtEndOfStream
		line = file.Readline
		line = TRIM(line)
	
		if UCase(line) = "<SCRIPT LANGUAGE=" & chr(34) & "VBSCRIPT" & chr(34) & ">" then
			dict.Add row, "!!||"
			row = row + 1
			abriu = 1
		end if
		
		if UCase(line) = "</SCRIPT>" then
			pode = 1
			abriu = 0
		end if
		
		if pode = 1 then
			dict.Add row, line
		end if
		
		if abriu = 1 then
			pode = 0
		end if
		
		row = row + 1
		
	Loop
	
	file.Close
	
	cont = 0
	Set arq = fso.CreateTextFile(project, True)
	For Each line in dict.Items
	
	if line <> "" then
		if InStr(line, "<script language="&chr(34)&"VBScript.Encode"&chr(34)&" src="&chr(34)& "Encoded/" & projectEn &chr(34)&"/>") <> 0 then
			line = Replace(line, "<script language="&chr(34)&"VBScript.Encode"&chr(34)&" src="&chr(34)& "Encoded/" & projectEn &chr(34)&"/>", "")
		end if
		
		if InStr(line, "!!||") <> 0 then
			line = Replace(line, "!!||", "<script language="&chr(34)&"VBScript.Encode"&chr(34)&" src="&chr(34)& "Encoded/" & projectEn &chr(34)&"/>")
		end if
		
		if InStr(line, "<script language="&chr(34)&"VBScript.Encode"&chr(34)&" src="&chr(34)& "Encoded/" &chr(34)&"/>") <> 0 then
			line = Replace(line, "<script language="&chr(34)&"VBScript.Encode"&chr(34)&" src="&chr(34)& "Encoded/" &chr(34)&"/>", "")
		end if
		
		if InStr(line, "<script language="&chr(34)&"VBScript.Encode"&chr(34)&" src="&chr(34)&""&chr(34)&"/>") <> 0 then
			line = Replace(line, "<script language="&chr(34)&"VBScript.Encode"&chr(34)&" src="&chr(34)&""&chr(34)&"/>", "")
		end if
		
		arq.WriteLine(line)
	end if
		
		cont = cont + 1
	Next
	
	arq.Close
	
	
End Function

'encoder("olhaele.vbs")
'extractVBS("olhaele.wsf")
'encoder("encoded/olhaele.vbs")
'injectVBE("olhaele.wsf")