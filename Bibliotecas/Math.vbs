'        @PARANAUERJ DEVELOPEMENT
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







Function Somar(a,b)
	if varType(a) < 8192 and varType(b) < 8192 then
		c = CDBL(a) + CDBL(b)
		Somar = c
	elseif (varType(a) < 8192 and varType(b) >= 8192) or (varType(a) >= 8192 and varType(b) < 8192) then
		Somar = "Entrada Invalida"

	elseif GetArrayDimAux(a) <> GetArrayDimAux(b) then
		Somar = "Entrada Invalida"

	elseif GetArrayDimAux(a) = 1 then
		Somar = sumVector(a,b)

	elseif GetArrayDimAux(a) = 2 then
		Somar = sumMatrix(a,b)
		
	end if
	end function


Function GetArrayDimAux(ByVal arr)
  GetArrayDim = Null
  Dim i
  If IsArray(arr) Then
    For i = 1 To 60
      On Error Resume Next
      UBound arr, i
      If Err.Number <> 0 Then
        GetArrayDim = i-1
        Exit Function
      End If
    Next
    GetArrayDim = i
  End If
End Function


Function sumVector(a,b)

	Redim vetorResp(0)

	if UBound(a) > UBound(b) then
		Redim vetorResp(UBound(a))
		menor = UBound(b)
		maior = UBound(a)
		mmaior = 1

	elseif UBound(a) = UBound(b) then
		Redim vetorResp(UBound(a))
		menor = UBound(b)
		maior = UBound(a)
		mmaior = 1

	elseif UBound(a) < UBound(b) then
		Redim vetorResp(UBound(b))
		menor = UBound(a)
		maior = UBound(b)
		mmaior = 2
	end if

	cont = 0

	while cont <= maior

		if cont <= menor then
			vetorResp(cont) = a(cont) + b(cont)
		else
			if mmaior = 1 then
				vetorResp(cont) = a(cont)
			else
				vetorResp(cont) = b(cont)
			end if
		end if

		cont = cont + 1

	wend

	sumVector = vetorResp

end function


function lineMatrix(a,n)

	cont = 0
	redim retorno(UBound(a,2))
	while cont <= UBound(a,1) 

		if cont = n then

			x = 0
			while x < UBound(a,2)
				retorno(x) = a(cont, x)
				x = x + 1
			wend

		end if
		
		cont = cont + 1

	wend

	lineMatrix = retorno

end function


Function sumMatrix(a,b)

	Redim vetorResp(0,0)

	if UBound(a) >= UBound(b) then
		if UBound(a,2) >= UBound(b,2) then
			Redim vetorResp(UBound(a),UBound(a,2))
			menor2 = UBound(b,2)
			maior2 = UBound(a,2)
			mmaior2 = 1

		elseif UBound(a,2) < UBound(b,2) then
			Redim vetorResp(UBound(a),UBound(b,2))
			menor2 = UBound(a,2)
			maior2 = UBound(b,2)
			mmaior2 = 2
		end if
		
		menor = UBound(b)
		maior = UBound(a)
		mmaior = 1

	elseif UBound(a) < UBound(b) then
		if UBound(a,2) >= UBound(b,2) then
			Redim vetorResp(UBound(b),UBound(a,2))
			menor2 = UBound(b,2)
			maior2 = UBound(a,2)
			mmaior2 = 1

		elseif UBound(a,2) < UBound(b,2) then
			Redim vetorResp(UBound(b),UBound(b,2))
			menor2 = UBound(a,2)
			maior2 = UBound(b,2)
			mmaior2 = 2
		end if
		menor = UBound(a)
		maior = UBound(b)
		mmaior = 2
	end if

	

	cont = 0
	
	'msgbox UBound(
	
	while cont <= maior
		y = 0

		if cont <= menor then
			while y <= maior2
				vetorResp(cont, y) = sumVector(lineMatrix(a, cont), lineMatrix(b, cont))(y)
				y = y + 1
			wend
		else
			if mmaior = 1 then
				while y <= maior2
					vetorResp(cont, y) = a(cont, y)
					y = y + 1
				wend
			else
				while y <= maior2
					vetorResp(cont, y) = b(cont, y)
					y = y + 1
				wend
			end if
		end if

		cont = cont + 1

	wend

	sumMatrix = vetorResp

end function



Function Subtrair(a,b)
	c = CDBL(a) - CDBL(b)
	Subtrair = c
	end function


Function Multiplicar(a,b)
	c = CDBL(a) * CDBL(b)
	Multiplicar = c
	end function


Function Dividir(a,b)
	if b = 0 then
		c = "Denominador igual a zero!"
		Dividir = c
	else 
		c = CDBL(a) / CDBL(b)
		Dividir = c
		end if
	end function


Function Exponenciacao(a,b)
	c = CDBL(a) ^ CDBL(b)
	Exponenciacao = c
	end function


Function Raiz(a,b)
	c = CDBL(a) ^ (1/CDBL(b))
	Raiz = c
	end function
	

Function Rest(a,b)
	c = a MOD b
	Rest = c
	end function


Function Delta(a,b,c)
	d = Exponenciacao(CDBL(b),2) - Multiplicar(Multiplicar(4,CDBL(a)),CDBL(c))
	Delta = d
	end function


Function Bhaskara(a,b,c)
	if Delta(CDBL(a),CDBL(b),CDBL(c)) >= 0 then
		Dim raizes(1)
		raizes(0) = (-CDBL(b) + Raiz(Delta(CDBL(a),CDBL(b),CDBL(c)),2))/(Multiplicar(2,CDBL(a)))
		raizes(1) = (-CDBL(b) - Raiz(Delta(CDBL(a),CDBL(b),CDBL(c)),2))/(Multiplicar(2,CDBL(a)))
		Bhaskara = raizes
	else
		Dim raizesa(1)
		raizesa(0) = "Nao tem resultado que pertenca aos reais"
		raizesa(1) = "Nao tem resultado que pertenca aos reais"
		Bhaskara = raizesa
		end if
	end function
	
	
Function DistanciaPontos(xa,ya,xb,yb)
	d = raiz(exponenciacao(CDBL(xa)-CDBL(xb),2) + exponenciacao(CDBL(ya)-CDBL(yb),2),2)
	DistanciaPontos = d
	end function
	

Function PI()

	Dim piE
	piE = 4 * Atn(1) 
	
	pi = piE

end function

Function Sec(X)

	Sec = 1 / Cos(X)

end function

Function Cosec(X)

	Cosec = 1 / Sin(X)

end function

Function Cotan(X)

	Cotan = 1 / Tan(X)

end function

Function Arcsin(X)

	Arcsin = Atn(X / Sqr(-X * X + 1))

end function

Function Arccos(X)

	Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)

end function


function det2d(matriz)

	determinante = (matriz(0,0) * matriz(1,1)) - (matriz(1,0) * matriz(0,1))
	
	det2d = determinante

end function

function det3d(matriz)

	'Coluna 1
	for i = 0 to 2
		redim matAux(2,2)
		if i = 0 then
			matAux(0,0) = matriz(1,1)
			matAux(0,1) = matriz(1,2)
			matAux(1,0) = matriz(2,1)
			matAux(1,1) = matriz(2,2)
			
		elseif i = 1 then
			matAux(0,0) = matriz(0,1)
			matAux(0,1) = matriz(0,2)
			matAux(1,0) = matriz(2,1)
			matAux(1,1) = matriz(2,2)
		else
			matAux(0,0) = matriz(0,1)
			matAux(0,1) = matriz(0,2)
			matAux(1,0) = matriz(1,1)
			matAux(1,1) = matriz(1,2)
		end if
		
		laplace = laplace + (matriz(i,0) * ((-1)^(i+1+1)) * det2d(matAux))
		
	next 
	
	det3d = laplace
	
end function

function det4d(matriz)

	'Coluna 1
	for i = 0 to 3
		redim matAux(3,3)
		
		if i = 0 then
			matAux(0,0) = matriz(1,1)
			matAux(0,1) = matriz(1,2)
			matAux(0,2) = matriz(1,3)
			matAux(1,0) = matriz(2,1)
			matAux(1,1) = matriz(2,2)
			matAux(1,2) = matriz(2,3)
			matAux(2,0) = matriz(3,1)
			matAux(2,1) = matriz(3,2)
			matAux(2,2) = matriz(3,3)
			
		elseif i = 1 then
			matAux(0,0) = matriz(0,1)
			matAux(0,1) = matriz(0,2)
			matAux(0,2) = matriz(0,3)
			matAux(1,0) = matriz(2,1)
			matAux(1,1) = matriz(2,2)
			matAux(1,2) = matriz(2,3)
			matAux(2,0) = matriz(3,1)
			matAux(2,1) = matriz(3,2)
			matAux(2,2) = matriz(3,3)
			
		elseif i = 2 then
			matAux(0,0) = matriz(0,1)
			matAux(0,1) = matriz(0,2)
			matAux(0,2) = matriz(0,3)
			matAux(1,0) = matriz(1,1)
			matAux(1,1) = matriz(1,2)
			matAux(1,2) = matriz(1,3)
			matAux(2,0) = matriz(3,1)
			matAux(2,1) = matriz(3,2)
			matAux(2,2) = matriz(3,3)
			
		else
			matAux(0,0) = matriz(0,1)
			matAux(0,1) = matriz(0,2)
			matAux(0,2) = matriz(0,3)
			matAux(1,0) = matriz(1,1)
			matAux(1,1) = matriz(1,2)
			matAux(1,2) = matriz(1,3)
			matAux(2,0) = matriz(2,1)
			matAux(2,1) = matriz(2,2)
			matAux(2,2) = matriz(2,3)
		end if
		
		laplace = laplace + (matriz(i,0) * ((-1)^(i+1+1)) * det3d(matAux))
		
	next 
	
	det4d = laplace
	
end function


function det5d(matriz)

	'Coluna 1
	for i = 0 to 3
		redim matAux(4,4)
		
		if i = 0 then
			matAux(0,0) = matriz(1,1)
			matAux(0,1) = matriz(1,2)
			matAux(0,2) = matriz(1,3)
			matAux(0,3) = matriz(1,4)
			matAux(1,0) = matriz(2,1)
			matAux(1,1) = matriz(2,2)
			matAux(1,2) = matriz(2,3)
			matAux(1,3) = matriz(2,4)
			matAux(2,0) = matriz(3,1)
			matAux(2,1) = matriz(3,2)
			matAux(2,2) = matriz(3,3)
			matAux(2,3) = matriz(3,4)
			matAux(3,0) = matriz(4,1)
			matAux(3,1) = matriz(4,2)
			matAux(3,2) = matriz(4,3)
			matAux(3,3) = matriz(4,4)
			
		elseif i = 1 then
			matAux(0,0) = matriz(0,1)
			matAux(0,1) = matriz(0,2)
			matAux(0,2) = matriz(0,3)
			matAux(0,3) = matriz(0,4)
			matAux(1,0) = matriz(2,1)
			matAux(1,1) = matriz(2,2)
			matAux(1,2) = matriz(2,3)
			matAux(1,3) = matriz(2,4)
			matAux(2,0) = matriz(3,1)
			matAux(2,1) = matriz(3,2)
			matAux(2,2) = matriz(3,3)
			matAux(2,3) = matriz(3,4)
			matAux(3,0) = matriz(4,1)
			matAux(3,1) = matriz(4,2)
			matAux(3,2) = matriz(4,3)
			matAux(3,3) = matriz(4,4)
			
		elseif i = 2 then
			matAux(0,0) = matriz(0,1)
			matAux(0,1) = matriz(0,2)
			matAux(0,2) = matriz(0,3)
			matAux(0,3) = matriz(0,4)
			matAux(1,0) = matriz(1,1)
			matAux(1,1) = matriz(1,2)
			matAux(1,2) = matriz(1,3)
			matAux(1,3) = matriz(1,4)
			matAux(2,0) = matriz(3,1)
			matAux(2,1) = matriz(3,2)
			matAux(2,2) = matriz(3,3)
			matAux(2,3) = matriz(3,4)
			matAux(3,0) = matriz(4,1)
			matAux(3,1) = matriz(4,2)
			matAux(3,2) = matriz(4,3)
			matAux(3,3) = matriz(4,4)
			
		elseif i = 3 then
			matAux(0,0) = matriz(0,1)
			matAux(0,1) = matriz(0,2)
			matAux(0,2) = matriz(0,3)
			matAux(0,3) = matriz(0,4)
			matAux(1,0) = matriz(1,1)
			matAux(1,1) = matriz(1,2)
			matAux(1,2) = matriz(1,3)
			matAux(1,3) = matriz(1,4)
			matAux(2,0) = matriz(2,1)
			matAux(2,1) = matriz(2,2)
			matAux(2,2) = matriz(2,3)
			matAux(2,3) = matriz(2,4)
			matAux(3,0) = matriz(4,1)
			matAux(3,1) = matriz(4,2)
			matAux(3,2) = matriz(4,3)
			matAux(3,3) = matriz(4,4)
		
		else
			matAux(0,0) = matriz(0,1)
			matAux(0,1) = matriz(0,2)
			matAux(0,2) = matriz(0,3)
			matAux(0,3) = matriz(0,4)
			matAux(1,0) = matriz(1,1)
			matAux(1,1) = matriz(1,2)
			matAux(1,2) = matriz(1,3)
			matAux(1,3) = matriz(1,4)
			matAux(2,0) = matriz(2,1)
			matAux(2,1) = matriz(2,2)
			matAux(2,2) = matriz(2,3)
			matAux(2,3) = matriz(2,4)
			matAux(3,0) = matriz(3,1)
			matAux(3,1) = matriz(3,2)
			matAux(3,2) = matriz(3,3)
			matAux(3,3) = matriz(3,4)
		end if
		
		laplace = laplace + (matriz(i,0) * ((-1)^(i+1+1)) * det4d(matAux))
		
	next 
	
	det5d = laplace
	
end function

function det(matriz)
	
	if UBound(matriz,1) <> UBound(matriz,2) then
		det = false
	else
		if UBound(matriz) = 2 then
			det = det2d(matriz)
		elseif UBound(matriz) = 3 then
			det = det3d(matriz)
		elseif UBound(matriz) = 4 then
			det = det4d(matriz)
		elseif UBound(matriz) = 5 then
			det = det5d(matriz)
		else
			det = false
		end if
	end if
	
end function