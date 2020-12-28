Sub show_table()
    set rs=server.createobject("adodb.Recordset")
    rs.Open SQL,cnn
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
    rs.Close
    set rs = nothing
End Sub

'轉日期時間格式
Function FDT(DT)
	FDT = FormatDateTime(DT,vbShortDate) & " " & FormatDateTime(DT,vbShortTime) & ":" & right("0" & Second(DT),2)
End Function

'=====取得cols & rows
Sub get_data()
set rs=server.createobject("adodb.Recordset")
rs.Open SQL,cnn

For i = 0 to rs.Fields.Count-1
	If i > 0 Then cols = cols & ","
	cols = cols & rs(i).Name	
Next
While Not rs.EOF
	If rows <> "" Then rows = rows & ";"
	For i = 0 to rs.Fields.Count-1
		If i > 0 Then rows = rows & ","
		rows = rows & rs(i)
	Next
rs.MoveNext	' 移到下一筆
Wend
rs.Close
set rs = nothing
End Sub


item_str = ""
'====取得單一欄位的字串
Sub get_item()
	set rs=server.createobject("adodb.Recordset")
	rs.Open SQL,cnn
	If Not rs.EOF Then
	item_str = rs(0)
	rs.MoveNext
	While Not rs.EOF
		item_str = item_str & "," & rs(0)
    rs.MoveNext
    Wend
	End If
	rs.Close
	set rs = nothing
End Sub


Function get_color(status)
	color = "#CCFFFF"
	If status = "RUN" Then
		color = "#00FF00"
	ElseIf status = "IDLE" Then
		color = "#FFFF00"	
	ElseIf status = "DOWN" Then
		color = "#FF0000"
	ElseIf status = "DMQC" Then
		color = "#FF00FF"
	ElseIf status = "PM" Then
		color = "#00FFFF"
	ElseIf status = "TEST" Then
		color = "#CCFFFF"
	End If
	get_color = color
End Function
