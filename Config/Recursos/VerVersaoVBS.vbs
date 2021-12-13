'        @PARANAUERJ DEVELOPEMENT
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







Set fso = CreateObject("Scripting.FileSystemObject")
vbsversion = fso.GetFileVersion("c:\windows\system32\wscript.exe")
vbsversion = CDbl(vbsversion)
vbsversion = vbsversion/10000000000000

if vbsversion < 5 then
	msgbox "Versao do VBScript menor que 5, pode ser que algumas funcionalidades nao funcionem corretamente." & vbCrlf & "Versao do VBScript instalada: " & vbsversion
else 
	
end if