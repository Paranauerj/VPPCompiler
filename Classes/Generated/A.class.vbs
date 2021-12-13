Class TESTE1
PRIVATE NICK
PRIVATE SENHA
PRIVATE NOME
PRIVATE TELEFONE
PRIVATE ENDERECO
PUBLIC Function SETNOME(PARAM1) 
NOME = PARAM1 
End Function
PUBLIC Function SETNICK(PARAM1) 
NICK = PARAM1 
End Function
PUBLIC Function SETSENHA(PARAM1) 
SENHA = MD5(PARAM1) 
End Function
PUBLIC Function SETTELEFONE(PARAM1) 
TELEFONE = PARAM1 
End Function
PUBLIC Function SETENDERECO(PARAM1) 
ENDERECO = PARAM1 
End Function
PUBLIC Function SAVEVALUES() 
AddValues ("BD.LIB:ADDVALUES USUARIO " & NICK & "|" & SENHA & "|" & ENDERECO & "|" & NOME & "|" & TELEFONE & "")
End Function
PUBLIC Function GETNICK() 
GETNICK = NICK
End Function
PUBLIC Function GETNOME() 
GETNOME = NOME
End Function
PUBLIC Function PEGABD(BASE, ROW) 
NICKS = GetValueRow(BASE, ROW)
PEGABD = NICKS
End Function
PUBLIC Function CHAMABD(BASE) 
FSO.CopyFile Comando.CurrentDirectory & "\Database\" & BASE & ".db", Comando.CurrentDirectory & "\Projetos\" & a & "\Database\"
End Function
Private Sub Class_Initialize(  )
End Sub
End Class
