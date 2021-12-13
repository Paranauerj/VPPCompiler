Class trabalhador
private nome
private idade
public FUNCTION  getInfo() 

info = "Idade: " & idade & vbcrlf & "Nome: " & nome 
getInfo = info
END FUNCTION
public cnh
public Function getInfo() 
info = getInfo() & vbcrlf & "cnh: " & cnh 
getInfo = info
End Function
Function GETnome()
GETnome = nome
End Function
Function GETidade()
GETidade = idade
End Function
Function GETcnh()
GETcnh = cnh
End Function
Function SETnome(param1)
nome = param1
End Function
Function SETidade(param1)
idade = param1
End Function
Function SETcnh(param1)
cnh = param1
End Function
Private Sub Class_Initialize(  )
End Sub
End Class
