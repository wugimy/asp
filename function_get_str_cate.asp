Function get_str_cate(S)
	ALL_NUMBER = True
	ALL_ENGLISH = True
	HAVE_BLANK = False
	'=====把字串正規化(英文為A，數字為0，其餘為#)=====
	FS = ""
	For k = 1 to len(S)
		C = Asc(Mid(S,k,1))
		If C >= 65 and C <= 91 Then
			FS = FS & "A"
		ElseIf C >= 48 and C <= 57 Then
			FS = FS & "0"
		Else
			FS = FS & "#"
		End If
	Next
	
	For k = 1 to Len(S)
		If Asc(Mid(S,k,1)) >= 48 And Asc(Mid(S,k,1)) <= 57 Then
			ALL_ENGLISH = False
		Else
			ALL_NUMBER = False
		End If
		If Mid(S,k,1) = " " Then
			HAVE_BLANK = True
		End If
	Next

	If len(S) = 7 and ALL_NUMBER = True Then
		S_TYPE = "工號"
	ElseIf FS = "AA0AA0" or FS = "AA0A00" Then
		S_TYPE = "DEPT"
	ElseIf instr(",CLA,SPT,CVD,IEX,RIE,PEN,WMA,HDP,WTO,MOR,MSH,LSR,LCV,","," & left(S,3) & ",") Then
		S_TYPE = "機台"
	ElseIf Left(S,2) = "A8" or Left(S,2) = "B8" Then
		S_TYPE = "機台"
	ElseIf (Asc(Left(S,1)) < 0 Or Asc(Left(S,1)) > 256) and len(S) >= 2 Then
		S_TYPE = "中文姓名"
	ElseIf ALL_ENGLISH = True and len(S) >= 3 Then
		S_TYPE = "英文姓名"
	ElseIf len(S) = 6 And (left(S,2)="AA" or left(S,2)="BP" or left(S,2)="BK" or left(S,2)="BB" or left(S,2)="CC" or left(S,2)="TT" or left(S,2)="AC" or left(S,2)="AB") and ALL_ENGLISH = False and ALL_NUMBER = False Then
		S_TYPE = "CRR_ID"
	ElseIf len(S) = 10 and IsNumeric(right(S,1)) and HAVE_BLANK = False Then
		S_TYPE = "LOT_ID"
	ElseIf len(S) = 9 and IsNumeric(right(S,1))=false and HAVE_BLANK = False Then
		S_TYPE = "SHT_ID"
	ElseIf len(S) = 2 Then
		S_TYPE = "EC_CODE"
	Else
		S_TYPE = "無法判定"
	End If
	
	get_str_cate = S_TYPE
End Function
