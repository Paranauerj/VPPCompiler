Class RELOGIO
PRIVATE SEGUNDO
PRIVATE MINUTO
PRIVATE HORA
Function GETSEGUNDO()
GETSEGUNDO = SEGUNDO
End Function
Function GETMINUTO()
GETMINUTO = MINUTO
End Function
Function GETHORA()
GETHORA = HORA
End Function
Function SETSEGUNDO(param1)
SEGUNDO = param1
End Function
Function SETMINUTO(param1)
MINUTO = param1
End Function
Function SETHORA(param1)
HORA = param1
End Function
PUBLIC Function INIT(SEGUNDOA, MINUTOA, HORAA) 
SEGUNDO = SEGUNDOA 
MINUTO = MINUTOA 
HORA = HORAA 
End Function
PUBLIC Function EQUALIZE(A) 
SEGUNDO = A.GETSEGUNDO() 
MINUTO = A.GETMINUTO 
HORA = A.GETHORA() 
End Function
Private Sub Class_Initialize(  )
End Sub
End Class
