Class carro
private id
public marca
public modelo
public ano
public preco
private base
private tem
Public Function Initialize(idObj, baseObj) 
id = idObj 
id = convert(id, "str") 
base = baseObj 
ids = GetValueRow(base, 0)
marcas = GetValueRow(base, 1)
modelos = GetValueRow(base, 2)
anos = GetValueRow(base, 3)
precos = GetValueRow(base, 4)
indice = InArray(ids, id) 
if indice <> -1  then

marca = marcas(indice) 
modelo = modelos(indice) 
ano = anos(indice) 
preco = precos(indice) 
tem = 1 
else
Message "ID: " & idObj & " nao foi encontrado!" & vbcrlf & "Tente Novamente!"
tem = 0 
end if
End Function
Public Function createNew(marcaObj, modeloObj, anoObj, precoObj) 
marca = marcaObj 
modelo = modeloObj 
ano = anoObj 
preco = precoObj 
tem = 1 
AddValues ("BD.Lib:AddValues " & base & " " & id & "|" & marcaObj & "|" & modeloObj & "|" & anoObj & "|" & precoObj & "")
End Function
Public Function isCar() 
isCar = tem
End Function
Function GETid()
GETid = id
End Function
Function GETmarca()
GETmarca = marca
End Function
Function GETmodelo()
GETmodelo = modelo
End Function
Function GETano()
GETano = ano
End Function
Function GETpreco()
GETpreco = preco
End Function
Function GETbase()
GETbase = base
End Function
Function GETtem()
GETtem = tem
End Function
Private Sub Class_Initialize(  )
End Sub
End Class
