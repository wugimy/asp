Sub show_table()
  Response.Write "<table><tr>"
  For i=0 to rs.Fields.Count-1
	  Response.Write "<th>" & rs(i).Name & "</th>"
  Next
  Response.Write "</tr>"
  While Not rs.EOF	' 判斷是否過了最後一筆
	  Response.Write "<tr>"
  	For i = 0 to rs.Fields.Count-1
	  	Response.Write "<TD>" & rs(i) & "</TD>"
	  Next
	  Response.Write "</tr>"
	  For i = 0 to UBound(S)
	  	If rs("PROD") <> "BS" Then
	  		S(i) = S(i) + rs("M" & i)
	  	End If
	  Next
  rs.MoveNext	' 移到下一筆
  Wend
End Sub
