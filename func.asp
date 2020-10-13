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
    rs.MoveNext	' 移到下一筆
    Wend
    Response.Write "</table>"
End Sub

'轉日期時間格式
Function FDT(DT)
	FDT = FormatDateTime(DT,vbShortDate) & " " & FormatDateTime(DT,vbShortTime) & ":" & right("0" & Second(DT),2)
End Function
