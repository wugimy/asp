<%
set fs = CreateObject("Scripting.FileSystemObject")
File = Server.MapPath("L5A公告.txt")
set txt = fs.OpenTextFile(File)

If Not fs.FileExists(File) Then
	Response.Write File & " 不存在!"
End If

If Not txt.atEndOfStream Then
	Content = txt.ReadAll
	Lines = Replace(Content,vbCrLf,"<br>")
	Response.Write Lines
End If
set fs = nothing
%>
