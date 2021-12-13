'        @PARANAUERJ DEVELOPEMENT op
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







set FSO = CreateObject("Scripting.FileSystemObject")
set WShell = WScript.CreateObject("WScript.Shell")
set Comando = WScript.CreateObject("WScript.Shell")


Function Bubble(array, ordem)
	
	if UCase(ordem) = "DESC" then
		n = UBound(array)
		Do
		  nn = -1
		  For j = LBound(array) to n - 1
			  If array(j) < array(j + 1) Then
				 TempValue = array(j + 1)
				 array(j + 1) = array(j)
				 array(j) = TempValue
				 nn = j
			  End If
		  Next
		  n = nn
		Loop Until nn = -1
		 
		s = ""
		For i = LBound(array) To UBound(array)
			s = s & array(i) & ","
		Next    
	
	else
		n = UBound(array)
		Do
		  nn = -1
		  For j = LBound(array) to n - 1
			  If array(j) > array(j + 1) Then
				 TempValue = array(j + 1)
				 array(j + 1) = array(j)
				 array(j) = TempValue
				 nn = j
			  End If
		  Next
		  n = nn
		Loop Until nn = -1
		 
		s = ""
		For i = LBound(array) To UBound(array)
			s = s & array(i) & ","
		Next  
	end if
	
End Function


Function inArray(arr, obj)
  ' On Error Resume Next
	x = 0
	
	indice = -1

    While x <= UBound(arr)
      If arr(x) = obj Then
        indice = x
		x = UBound(arr) + 1
      End If
	  x = x + 1
    Wend
	
  'Err.Clear()
  inArray = indice

End Function


Function addDict(arr, key, value)

	retorno = true

	if isset(key) then
		call arr.add(key, value)
		retorno = true
	else
		retorno = false
	end if
	
	addDict = retorno

End Function


Function push(arr, value)

	chaves = arr.Keys
	key = UBound(chaves) + 1

	call arr.add(key, value)

End Function


Function pop(arr)

	retorno = -1

	chaves = arr.Keys
	itens = arr.Items	

	retorno = arr(chaves(UBound(chaves)))
	arr.Remove chaves(UBound(chaves))

	pop = retorno

End Function



Function ReSize(arr, num)

	ReDim PRESERVE arr(num)
	

End Function


Function sizeOf(arr)

	sizeOf = UBound(arr)

end Function

Function GetArrayDim(ByVal arr)
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