<%
FullPath = Request.ServerVariables("PATH_TRANSLATED")
FullPath = Left(FullPath,InStrRev(FullPath,"\"))
'Response.Write  & "<br>"
'FullPath =Request.ServerVariables("APPL_PHYSICAL_PATH")
'Response.Write FullPath & "<br>"
Set fs = Server.CreateObject("Scripting.FileSystemObject")  
Set FileCollection = fs.GetFolder(FullPath)  
For Each file in FileCollection.files
If Left(file.Type,3) = "PDF" Then
Response.Write "<table width='100%'>" 
Response.Write "<td width='20%'>檔案名稱</td>" 
Response.Write "<td width='30%'>檔案類型</td>" 
Response.Write "<td width='25%'>建立日期</td>" 
Response.Write "<td width='25%'>存取日期</td></tr>" 
response.write "<td><a target=_blank href=" & file.name & ">" & file.name & "</a></td>" 
response.write "<td>"&file.Type&"</td>"   
response.write "<td>"&file.DateCreated&"</td>" 
response.write "<td>"&file.DateLastAccessed&"</td></tr>"
End If
Next 
response.write "</table>" 
%> 
