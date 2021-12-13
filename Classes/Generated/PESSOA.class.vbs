Class pessoa
private nome
private idade
public Function getInfo() 
info = "Idade: " & idade & vbcrlf & "Nome: " & nome 
getInfo = info
End Function
Function GETnome()
GETnome = nome
End Function
Function GETidade()
GETidade = idade
End Function
Function SETnome(param1)
nome = param1
End Function
Function SETidade(param1)
idade = param1
End Function
Private Sub Class_Initialize(  )
End Sub
End Class
