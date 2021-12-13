'        @PARANAUERJ DEVELOPEMENT UPDATED 2
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







Function GetJSONDataStr(file)
Set FSO = CreateObject("Scripting.FileSystemObject")
Set arq = FSO.OpenTextFile(file, 1, true)
linhas = 0
co = 0
Do While arq.AtEndOfStream <> True
   arq.SkipLine
   linhas = linhas + 1
Loop
'msgbox linhas
arq.close
ReDim Buf(linhas)

Set arq = FSO.OpenTextFile(file, 1, True)

while co < linhas

	Buf(co) = arq.ReadLine()
	'msgbox Buf(co)
	co = co + 1

wend
arq.close

GetJSONDataStr = Join(Buf, vbCrLf)
End Function


Class JSONStringEncoder

    Private m_RegExp
    
    Sub Class_Initialize()
        Set m_RegExp = Nothing
    End Sub
    
    Function Encode(ByVal Str)

        Dim Parts(): ReDim Parts(3)
        Dim NextPartIndex: NextPartIndex = 0
        Dim AnchorIndex: AnchorIndex = 1
        Dim CharCode, Escaped
        Dim Match, MatchIndex
        Dim RegExp: Set RegExp = m_RegExp
        If RegExp Is Nothing Then
            Set RegExp = New RegExp
            ' See https://github.com/douglascrockford/JSON-js/blob/43d7836c8ec9b31a02a31ae0c400bdae04d3650d/json2.js#L196
            RegExp.Pattern = "[\\\""\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]"
            RegExp.Global = True
            Set m_RegExp = RegExp
        End If
        For Each Match In RegExp.Execute(Str)
            MatchIndex = Match.FirstIndex + 1
            If NextPartIndex > UBound(Parts) Then ReDim Preserve Parts(UBound(Parts) * 2)
            Parts(NextPartIndex) = Mid(Str, AnchorIndex, MatchIndex - AnchorIndex): NextPartIndex = NextPartIndex + 1
            CharCode = AscW(Mid(Str, MatchIndex, 1))
            Select Case CharCode
                Case 34  : Escaped = "\"""
                Case 10  : Escaped = "\n"
                Case 13  : Escaped = "\r"
                Case 92  : Escaped = "\\"
                Case 8   : Escaped = "\b"
                Case Else: Escaped = "\u" & Right("0000" & Hex(CharCode), 4)
            End Select
            If NextPartIndex > UBound(Parts) Then ReDim Preserve Parts(UBound(Parts) * 2)
            Parts(NextPartIndex) = Escaped: NextPartIndex = NextPartIndex + 1
            AnchorIndex = MatchIndex + 1
        Next
        If AnchorIndex = 1 Then Encode = """" & Str & """": Exit Function
        If NextPartIndex > UBound(Parts) Then ReDim Preserve Parts(UBound(Parts) * 2)
        Parts(NextPartIndex) = Mid(Str, AnchorIndex): NextPartIndex = NextPartIndex + 1
        ReDim Preserve Parts(NextPartIndex - 1)
        Encode = """" & Join(Parts, "") & """"
    
    End Function

End Class

Dim TheJSONStringEncoder: Set TheJSONStringEncoder = New JSONStringEncoder

Function EncodeJSONString(ByVal Str)
    EncodeJSONString = TheJSONStringEncoder.Encode(Str) 
End Function

Function EncodeJSONMember(ByVal Key, Value)
    EncodeJSONMember = EncodeJSONString(Key) & ":" & JSONStringify(Value)
End Function

Public Function JSONStringify(Thing) 

    Dim Key, Item, Index, NextIndex, Arr()
    Dim VarKind: VarKind = VarType(Thing)
    Select Case VarKind
        Case vbNull, vbEmpty: JSONStringify = "null"
        Case vbDate: JSONStringify = EncodeJSONString(FormatISODateTime(Thing))
        Case vbString: JSONStringify = EncodeJSONString(Thing)
        Case vbBoolean: If Thing Then JSONStringify = "true" Else JSONStringify = "false"
        Case vbObject
            If Thing Is Nothing Then
                JSONStringify = "null"
            Else
                If TypeName(Thing) = "Dictionary" Then
                    If Thing.Count = 0 Then JSONStringify = "{}": Exit Function
                    ReDim Arr(Thing.Count - 1)
                    Index = 0
                    For Each Key In Thing.Keys
                        Arr(Index) = EncodeJSONMember(Key, Thing(Key))
                        Index = Index + 1
                    Next
                    JSONStringify = "{" & Join(Arr, ",") & "}"
                Else
                    ReDim Arr(3)
                    NextIndex = 0
                    For Each Item In Thing
                        If NextIndex > UBound(Arr) Then ReDim Preserve Arr(UBound(Arr) * 2)
                        Arr(NextIndex) = JSONStringify(Item)
                        NextIndex = NextIndex + 1
                    Next
                    ReDim Preserve Arr(NextIndex - 1)
                    JSONStringify = "[" & Join(Arr, ",") & "]"
                End If
            End If
        Case Else
            If vbArray = (VarKind And vbArray) Then
                For Index = LBound(Thing) To UBound(Thing)
                    If Len(JSONStringify) > 0 Then JSONStringify = JSONStringify & ","
                    JSONStringify = JSONStringify & JSONStringify(Thing(Index))
                Next
                JSONStringify = "[" & JSONStringify & "]"
            ElseIf IsNumeric(Thing) Then
                JSONStringify = CStr(Thing)
            Else
                JSONStringify = EncodeJSONString(CStr(Thing))
            End If
    End Select

End Function

function jsonEncode(thing)
	retorno = JSONStringify(thing)
	retorno = "[" & retorno & "]"
	jsonEncode = retorno
end function