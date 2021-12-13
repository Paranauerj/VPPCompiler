'        @PARANAUERJ DEVELOPEMENT UPDATED 2
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







set FSO = CreateObject("Scripting.FileSystemObject")
set WShell = WScript.CreateObject("WScript.Shell")
set Comando = WScript.CreateObject("WScript.Shell")


Function CriaBD(a)
	on error resume next
	if NOT FSO.FolderExists(WShell.CurrentDirectory & "\Database") then
		set arq = FSO.CreateFolder (WShell.CurrentDirectory & "\Database")
		WScript.Sleep 1500
		end if
		
	if NOT FSO.FileExists (WShell.CurrentDirectory & "\Database\" & a & ".db") then 
		set arq = FSO.CreateTextFile(WShell.CurrentDirectory & "\Database\" & a & ".db", 8)
			arq.write("ROWS///")
		arq.Close
	else
		msgbox "Database ja existente"
		end if
		
	End Function
	
	

Function CsvToVppBuild(a)
	on error resume next
	if NOT FSO.FolderExists(WShell.CurrentDirectory & "\Database") then
		set arq = FSO.CreateFolder (WShell.CurrentDirectory & "\Database")
		WScript.Sleep 1500
		end if
		
	if FSO.FileExists (WShell.CurrentDirectory & "\Database\" & a & ".csv") then 
		set arq = FSO.CreateTextFile(WShell.CurrentDirectory & "\Database\" & a & ".db", 8)
			if Lines(a, "csv") >= 32000 then
				msgbox "Aviso! Numero de linhas excede o maximo permitido!"
			end if
			'arq.write("ROWS///")
		arq.Close
	else
		msgbox "Database com nome " & a & " ja existente!"
		end if
		
	End Function
	
	

Function VppToCsvBuild(a)
	on error resume next
	if NOT FSO.FolderExists(WShell.CurrentDirectory & "\Database") then
		set arq = FSO.CreateFolder (WShell.CurrentDirectory & "\Database")
		WScript.Sleep 1500
		end if
		
	if FSO.FileExists (WShell.CurrentDirectory & "\Database\" & a & ".db") then 
		set arq = FSO.CreateTextFile(WShell.CurrentDirectory & "\Database\" & a & ".csv", 8)
			if Lines(a, "db") >= 32000 then
				msgbox "Aviso! Numero de linhas excede o maximo permitido!"
			end if
			'arq.write("ROWS///")
		arq.Close
	else
		msgbox "Database com nome " & a & " ja existente!"
		end if
		
	End Function


Function VppToCsvConvert(a)
	on error resume next
		
	if FSO.FileExists (WShell.CurrentDirectory & "\Database\" & a & ".csv") and FSO.FileExists (WShell.CurrentDirectory & "\Database\" & a & ".db") then 
		Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".db", 1)
		Set arq2 = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".csv", 2)
		
		cont = 0
		
		do while not arq.AtEndOfStream
		
		linha = arq.ReadLine
		
			if cont = 0 then
				linha = replace(linha, ",", ".")
				linha = replace(linha, "/!", "///")
				linha = replace(linha, "ROWS///", "")
				linha = replace(linha, "///", ",")
				arq2.WriteLine(linha)
				
				cont = 1
				
			else
				linha = replace(linha, ",", ".")
				linha = replace(linha, "/!", "///")
				linha = replace(linha, "INSERT///", "")
				linha = replace(linha, "///", ",")
				arq2.WriteLine(linha)
				
			end if
			
		Loop
		
		arq2.Close
		arq.Close
	else
		msgbox "Database com nome " & a & " nao existente!"
		end if
		
	End Function



Function Lines(a, ext)

		on error resume next
		
		Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & "." & ext & "", 1)
		
		cont = 1
		
		do while not arq.AtEndOfStream
			
			arq.SkipLine
			' cont = cont + 1
			
		Loop
		
		arq.Close
		
		cont = arq.Line-1
		
		Lines = cont

	End Function



Function CsvToVppConvert(a)
	on error resume next
		
	if FSO.FileExists (WShell.CurrentDirectory & "\Database\" & a & ".db") and FSO.FileExists (WShell.CurrentDirectory & "\Database\" & a & ".csv") then 
		Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".csv", 1)
		Set arq2 = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".db", 2)
		
		cont = 0
		
		do while not arq.AtEndOfStream
		
		linha = arq.ReadLine
		
			if cont = 0 then
				
				linha = replace(linha, "///", "/!")
				linha = replace(linha, ",", "///")
				arq2.WriteLine("ROWS///" & linha)
				
				cont = 1
				
			else
				linha = replace(linha, "///", "/!")
				linha = replace(linha, ",", "///")
				arq2.WriteLine("INSERT///" & linha)
				
			end if
			
		Loop
		
		arq2.Close
		arq.Close
	else
		msgbox "Database com nome " & a & " nao existente!"
		end if
		
	End Function
	
	
	
Function AddRow(a,b)

	If FSO.FileExists(WShell.CurrentDirectory & "\Database\" & a & ".db") then
		Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".db", 1)
		NumLinhas = 0
		kaka = 0
		okay = "0"
		do while not arq.AtEndOfStream
			haha = arq.ReadLine
			NumLinhas = NumLinhas + 1
		Loop
		arq.Close
		
		Set arq3 = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".db", 1)
		do while not arq3.AtEndOfStream
			haha = Split(TRIM(arq3.ReadLine), "///")
			if kaka = NumLinhas - 1 then
				if haha(0) = "ROWS" then
					okay = "1"
				end if
			end if
			kaka = kaka + 1
		Loop
		arq3.Close
		
		Set arq2 = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".db", 8)
		if okay = "1" then

			if haha(UBound(haha)) = "" then
			else
				arq2.Write("///")
			end if

			arq2.Write(b)

		end if
		if okay = "0" then
			msgbox "Tabela ja criada! Nao e possivel inserir rows na tabela!"
		end if
		arq2.Close
	else
		msgbox "Database " & a & " nao encontrada!"
		end if
	
	End Function
	
Function AiRow(a,b)

	If FSO.FileExists(WShell.CurrentDirectory & "\Database\" & a & ".db") then
		Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".db", 1)
		erro = 0
		kaka = 1
		do while not arq.AtEndOfStream
			haha = split(arq.ReadLine, "///")
			if haha(0) = "ROWS" then
				while kaka <= UBound(haha)
					if haha(kaka) = b then
						erro = 40
						end if
					kaka = kaka + 1
				wend
			end if
		Loop
		arq.Close
		if erro = 1 or erro = 0 then
			msgbox "Certifique-se de que o Row existe!"
		else		
			Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".db", 8)
			arq.WriteLine( vbCrlf & "AI " & b)
			arq.Close
		end if
	else
		msgbox "Database " & a & " nao encontrada!"
		end if
	
	
	End Function
	
	
Function AddValues(a)
	
	c = Split(TRIM(a), " ", 3)
	d = Split(TRIM(c(2)), "|")

	If FSO.FileExists(WShell.CurrentDirectory & "\Database\" & c(1) & ".db") then
		Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & c(1) & ".db", 1)
		kaka = 0
		Numrows = 0
		do while not arq.AtEndOfStream
			haha = split(arq.ReadLine, "///")
			if haha(0) = "ROWS" then
				Numrows = UBound(haha) - 1
				kaka = 1
			end if
		Loop
		arq.Close
		Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & c(1) & ".db", 8)
		alcance = UBound(d)
		if kaka = 1 and (Numrows) = (alcance) then
			arq.Write(vbCrlf & "INSERT")
			hihi = 0
			while hihi <= alcance
				arq.Write("///" & d(hihi))
				hihi = hihi + 1
			wend
		else
			msgbox "Tabela nao definida ou numero de rows errado!"
		end if
		arq.Close
	else
		msgbox "Tabela " & c(1) & " nao encontrada!"
	end if
	
	
	End function
	
	
Function GetRows(a)
	
	If FSO.FileExists(WShell.CurrentDirectory & "\Database\" & a & ".db") then
		Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".db", 1)
		rows = ""
		zum = 1
		do while not arq.AtEndOfStream
			haha = split(arq.ReadLine, "///")
			if haha(0) = "ROWS" then
				while zum <= UBound(haha)
					rows = rows & haha(zum) & "///"
					' rows = Replace(rows, "_", " ")
					zum = zum + 1
				wend
			end if
		Loop
		arq.Close
		
		else
			msgbox "Database " & a & " nao existe!"
			end if
	b = TRIM(rows)
	
	GetRows = b
	
	End Function
	
	
Function GetValues(a)

	If FSO.FileExists(WShell.CurrentDirectory & "\Database\" & a & ".db") then
		Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".db", 1)
		Set dicionario = CreateObject("Scripting.Dictionary")
		dicionario.RemoveAll
		hihihi = 0
		Do Until arq.AtEndOfStream
			novalinha = ""
			haha = split(arq.ReadLine, "///")
			if haha(0) = "INSERT" then
				pimbis = 1
				while pimbis <= UBound(haha)
					novalinha = novalinha & haha(pimbis) & "///"
					' novalinha = Replace(novalinha, "_", " ")
					pimbis = pimbis + 1
				wend
				novalinha = TRIM(novalinha)
				dicionario.Add hihihi, novalinha
				hihihi = hihihi + 1
			end if
		Loop
		arq.Close
	end if
	
	pimbamaster = ""
	For Each elem In dicionario
		pimbamaster = pimbamaster & dicionario(elem) & vbCrlf
	Next

	GetValues = pimbamaster
	
	End function
	
	
Function GetValueRow(a,b)
	
	If FSO.FileExists(WShell.CurrentDirectory & "\Database\" & a & ".db") then
		Set arq = FSO.OpenTextFile(Comando.CurrentDirectory & "\Database\" & a & ".db", 1)
		linhas = Split(TRIM(GetValues(a)), vbCrlf)
		RequestedRow = ""
		x = 0
		linha = Split(linhas(0), "///")
		bint = Cint(b)
		if bint <= UBound(linha) then
		while x < UBound(linhas)
			linha = Split(linhas(x), "///")
			RequestedRow = RequestedRow & "///" & linha(bint)
			RequestedRow = Replace(RequestedRow, "_", " ")
			x = x + 1
		wend
			else 
			msgbox "Row " & b & " invalido" & vbCrlf & "Rows validos nessa tabela vao de 0 a " & UBound(linhas) - 1
		end if
		arq.Close
		else
			msgbox "Tabela " & a & " nao existe!"
	end if
	
	RequestedRow = TRIM(RequestedRow)
	
	c = Split(RequestedRow, "///")
	
	GetValueRow = c

	End function
	


	Private Const BITS_TO_A_BYTE = 8
	Private Const BYTES_TO_A_WORD = 4
	Private Const BITS_TO_A_WORD = 32
	Private m_lOnBits(30)
	Private m_l2Power(30)

	m_lOnBits(0) = CLng(1)
	m_lOnBits(1) = CLng(3)
	m_lOnBits(2) = CLng(7)
	m_lOnBits(3) = CLng(15)
	m_lOnBits(4) = CLng(31)
	m_lOnBits(5) = CLng(63)
	m_lOnBits(6) = CLng(127)
	m_lOnBits(7) = CLng(255)
	m_lOnBits(8) = CLng(511)
	m_lOnBits(9) = CLng(1023)
	m_lOnBits(10) = CLng(2047)
	m_lOnBits(11) = CLng(4095)
	m_lOnBits(12) = CLng(8191)
	m_lOnBits(13) = CLng(16383)
	m_lOnBits(14) = CLng(32767)
	m_lOnBits(15) = CLng(65535)
	m_lOnBits(16) = CLng(131071)
	m_lOnBits(17) = CLng(262143)
	m_lOnBits(18) = CLng(524287)
	m_lOnBits(19) = CLng(1048575)
	m_lOnBits(20) = CLng(2097151)
	m_lOnBits(21) = CLng(4194303)
	m_lOnBits(22) = CLng(8388607)
	m_lOnBits(23) = CLng(16777215)
	m_lOnBits(24) = CLng(33554431)
	m_lOnBits(25) = CLng(67108863)
	m_lOnBits(26) = CLng(134217727)
	m_lOnBits(27) = CLng(268435455)
	m_lOnBits(28) = CLng(536870911)
	m_lOnBits(29) = CLng(1073741823)
	m_lOnBits(30) = CLng(2147483647)
	m_l2Power(0) = CLng(1)
	m_l2Power(1) = CLng(2)
	m_l2Power(2) = CLng(4)
	m_l2Power(3) = CLng(8)
	m_l2Power(4) = CLng(16)
	m_l2Power(5) = CLng(32)
	m_l2Power(6) = CLng(64)
	m_l2Power(7) = CLng(128)
	m_l2Power(8) = CLng(256)
	m_l2Power(9) = CLng(512)
	m_l2Power(10) = CLng(1024)
	m_l2Power(11) = CLng(2048)
	m_l2Power(12) = CLng(4096)
	m_l2Power(13) = CLng(8192)
	m_l2Power(14) = CLng(16384)
	m_l2Power(15) = CLng(32768)
	m_l2Power(16) = CLng(65536)
	m_l2Power(17) = CLng(131072)
	m_l2Power(18) = CLng(262144)
	m_l2Power(19) = CLng(524288)
	m_l2Power(20) = CLng(1048576)
	m_l2Power(21) = CLng(2097152)
	m_l2Power(22) = CLng(4194304)
	m_l2Power(23) = CLng(8388608)
	m_l2Power(24) = CLng(16777216)
	m_l2Power(25) = CLng(33554432)
	m_l2Power(26) = CLng(67108864)
	m_l2Power(27) = CLng(134217728)
	m_l2Power(28) = CLng(268435456)
	m_l2Power(29) = CLng(536870912)
	m_l2Power(30) = CLng(1073741824)

Private Function LShift(lValue, iShiftBits)
	
	If iShiftBits = 0 Then
		LShift = lValue
		
		Exit Function
		
	ElseIf iShiftBits = 31 Then
		
		If lValue And 1 Then
			LShift = &H80000000
		Else
			LShift = 0
		End If
		
		Exit Function
		
	ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
		Err.Raise 6
	End If
	
	If (lValue And m_l2Power(31 - iShiftBits)) Then
		LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
	Else
		LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
	End If
	
End Function

Private Function RShift(lValue, iShiftBits)
	
	If iShiftBits = 0 Then
		RShift = lValue
		
		Exit Function
		
	ElseIf iShiftBits = 31 Then
		If lValue And &H80000000 Then
			RShift = 1
		Else
			RShift = 0
		End If
		
		Exit Function
		
	ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
		Err.Raise 6
	End If
	
	RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
	
	If (lValue And &H80000000) Then
		RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
	End If
	
End Function

Private Function RotateLeft(lValue, iShiftBits)
	RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
Dim lX4
Dim lY4
Dim lX8
Dim lY8
Dim lResult

	lX8 = lX And &H80000000
	lY8 = lY And &H80000000
	lX4 = lX And &H40000000
	lY4 = lY And &H40000000
	lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

	If lX4 And lY4 Then
		lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
	ElseIf lX4 Or lY4 Then
		If lResult And &H40000000 Then
			lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
		Else
			lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
		End If
	Else
		lResult = lResult Xor lX8 Xor lY8
	End If
	
	AddUnsigned = lResult
	
End Function

Private Function F(x, y, z)
	F = (x And y) Or ((Not x) And z)
End Function

Private Function G(x, y, z)
	G = (x And z) Or (y And (Not z))
End Function

Private Function H(x, y, z)
	H = (x Xor y Xor z)
End Function

Private Function I(x, y, z)
	I = (y Xor (x Or (Not z)))
End Function

Private Sub FF(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac))
	a = RotateLeft(a, s)
	a = AddUnsigned(a, b)
End Sub

Private Sub GG(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac))
	a = RotateLeft(a, s)
	a = AddUnsigned(a, b)
End Sub

Private Sub HH(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac))
	a = RotateLeft(a, s)
	a = AddUnsigned(a, b)
End Sub

Private Sub II(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac))
	a = RotateLeft(a, s)
	a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(sMessage)
Dim lMessageLength
Dim lNumberOfWords
Dim lWordArray()
Dim lBytePosition
Dim lByteCount
Dim lWordCount
Const MODULUS_BITS = 512
Const CONGRUENT_BITS = 448

	lMessageLength = Len(sMessage)
	lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
	ReDim lWordArray(lNumberOfWords - 1)
	lBytePosition = 0
	lByteCount = 0

	Do Until lByteCount >= lMessageLength
		lWordCount = lByteCount \ BYTES_TO_A_WORD
		lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
		lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
		lByteCount = lByteCount + 1
	Loop
	
	lWordCount = lByteCount \ BYTES_TO_A_WORD
	lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
	lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
	lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
	lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
	
	ConvertToWordArray = lWordArray
	
End Function

Private Function WordToHex(lValue)
Dim lByte
Dim lCount

	For lCount = 0 To 3
		lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
		WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
	Next
	
End Function

Public Function MD5(sMessage)
Dim x
Dim k
Dim AA
Dim BB
Dim CC
Dim DD
Dim a
Dim b
Dim c
Dim d
Const S11 = 7
Const S12 = 12
Const S13 = 17
Const S14 = 22
Const S21 = 5
Const S22 = 9
Const S23 = 14
Const S24 = 20
Const S31 = 4
Const S32 = 11
Const S33 = 16
Const S34 = 23
Const S41 = 6
Const S42 = 10
Const S43 = 15
Const S44 = 21

	x = ConvertToWordArray(sMessage)
	a = &H67452301
	b = &HEFCDAB89
	c = &H98BADCFE
	d = &H10325476
	
	For k = 0 To UBound(x) Step 16
		AA = a
		BB = b
		CC = c
		DD = d
		FF a, b, c, d, x(k + 0), S11, &HD76AA478
		FF d, a, b, c, x(k + 1), S12, &HE8C7B756
		FF c, d, a, b, x(k + 2), S13, &H242070DB
		FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
		FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
		FF d, a, b, c, x(k + 5), S12, &H4787C62A
		FF c, d, a, b, x(k + 6), S13, &HA8304613
		FF b, c, d, a, x(k + 7), S14, &HFD469501
		FF a, b, c, d, x(k + 8), S11, &H698098D8
		FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
		FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
		FF b, c, d, a, x(k + 11), S14, &H895CD7BE
		FF a, b, c, d, x(k + 12), S11, &H6B901122
		FF d, a, b, c, x(k + 13), S12, &HFD987193
		FF c, d, a, b, x(k + 14), S13, &HA679438E
		FF b, c, d, a, x(k + 15), S14, &H49B40821
		GG a, b, c, d, x(k + 1), S21, &HF61E2562
		GG d, a, b, c, x(k + 6), S22, &HC040B340
		GG c, d, a, b, x(k + 11), S23, &H265E5A51
		GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
		GG a, b, c, d, x(k + 5), S21, &HD62F105D
		GG d, a, b, c, x(k + 10), S22, &H2441453
		GG c, d, a, b, x(k + 15), S23, &HD8A1E681
		GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
		GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
		GG d, a, b, c, x(k + 14), S22, &HC33707D6
		GG c, d, a, b, x(k + 3), S23, &HF4D50D87
		GG b, c, d, a, x(k + 8), S24, &H455A14ED
		GG a, b, c, d, x(k + 13), S21, &HA9E3E905
		GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
		GG c, d, a, b, x(k + 7), S23, &H676F02D9
		GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
		HH a, b, c, d, x(k + 5), S31, &HFFFA3942
		HH d, a, b, c, x(k + 8), S32, &H8771F681
		HH c, d, a, b, x(k + 11), S33, &H6D9D6122
		HH b, c, d, a, x(k + 14), S34, &HFDE5380C
		HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
		HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
		HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
		HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
		HH a, b, c, d, x(k + 13), S31, &H289B7EC6
		HH d, a, b, c, x(k + 0), S32, &HEAA127FA
		HH c, d, a, b, x(k + 3), S33, &HD4EF3085
		HH b, c, d, a, x(k + 6), S34, &H4881D05
		HH a, b, c, d, x(k + 9), S31, &HD9D4D039
		HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
		HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
		HH b, c, d, a, x(k + 2), S34, &HC4AC5665
		II a, b, c, d, x(k + 0), S41, &HF4292244
		II d, a, b, c, x(k + 7), S42, &H432AFF97
		II c, d, a, b, x(k + 14), S43, &HAB9423A7
		II b, c, d, a, x(k + 5), S44, &HFC93A039
		II a, b, c, d, x(k + 12), S41, &H655B59C3
		II d, a, b, c, x(k + 3), S42, &H8F0CCC92
		II c, d, a, b, x(k + 10), S43, &HFFEFF47D
		II b, c, d, a, x(k + 1), S44, &H85845DD1
		II a, b, c, d, x(k + 8), S41, &H6FA87E4F
		II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
		II c, d, a, b, x(k + 6), S43, &HA3014314
		II b, c, d, a, x(k + 13), S44, &H4E0811A1
		II a, b, c, d, x(k + 4), S41, &HF7537E82
		II d, a, b, c, x(k + 11), S42, &HBD3AF235
		II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
		II b, c, d, a, x(k + 9), S44, &HEB86D391
		a = AddUnsigned(a, AA)
		b = AddUnsigned(b, BB)
		c = AddUnsigned(c, CC)
		d = AddUnsigned(d, DD)
	Next

	MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
	
End Function

	
'Function DeleteBD(a)


	'End function