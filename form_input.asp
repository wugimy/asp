<form method="POST" action="<%=Request.ServerVariables("PATH_INFO")%>">
<input type="submit" value="送出" name="B1">
</form>

<table>
<%
For i=0 to rs.Fields.Count-1
%>
<tr><td>
<%=rs(i).Name%>:</td><td><input type="text" name="<%=rs(i).Name%>" value="<%=field_value(i)%>"><br>
<input type="hidden" name="FN" value="<%=rs(i).Name%>">
</td></tr>
<%
Next
%>
</table>
