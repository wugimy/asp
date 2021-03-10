'=====================================
'以下為計算機的功能
Function calculate(S)
	RESULT = ""
	N = ""
	OP = ""
	For i = 1 to len(S)
		C = Mid(S,i,1)
		If C = "+" or C = "-" or C = "*" or C = "/" Then
			If OP = "" Then
				If IsNumeric(N) Then
					RESULT = CDBL(N)
				Else
					RESULT = CDBL(0)
				End If
			Else
				If IsNumeric(N) Then
					RESULT = get_result(RESULT,N,OP)
				End If
			End If
			OP = C
			N = ""
		ElseIf C = "=" Then
			RESULT = get_result(RESULT,N,OP)
		Else
			N = N & C
		End If
	Next
	calculate = RESULT
End Function

Function get_result(N0,N1,OP)
	If OP = "+" Then
		RESULT = N0 + N1
	ElseIf OP = "-" Then
		RESULT = N0 - N1
	ElseIf OP = "*" Then
		RESULT = N0 * N1
	ElseIf OP = "/" Then
		RESULT = N0 / N1
	End If
	get_result = RESULT
End Function
'
'=====================================
