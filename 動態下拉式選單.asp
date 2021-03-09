<!DOCTYPE html>
<HTML>

<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>動態下拉式選單 (二階層)</Title>

<script type="text/javascript" src="/js/jquery-1.6.min.js"></script>

</Head>

<body>

<%
Set cnn = Server.CreateObject("ADODB.Connection")
cnn.Open "Driver={MySQL ODBC 5.1 Driver};server=l5abweb01;UID=Mysql;PWD=Mysql;database=l5ab_db;Option=4;"
SQL = "select ABBR_NO,max(LOT_TYPE) as LOT_TYPE,max(PRODUCT_CODE) as PRODUCT_CODE from c_rou_prod where SITE_ID='A' and LOT_TYPE='PROD' group by ABBR_NO"
SQL = "select PRODUCT_CODE,group_concat(ABBR_NO) from (" & SQL & ") A group by PRODUCT_CODE"
'Call show_table()

    set rs=server.createobject("adodb.Recordset")
	rs.CursorLocation = 3
	rs.Open SQL,cnn
	ReDim model(rs.RecordCount-1)
	ReDim model_ec(UBound(model))
	For i = 0 to UBound(model)
		model(i) = rs(0)
		model_ec(i) = rs(1)
		rs.MoveNext
	Next
    rs.Close
    set rs = nothing

cnn.Close
set cnn = nothing

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

Response.Write join(model,",")
%>

<form name="myForm" method="POST" action="<%=Request.ServerVariables("PATH_INFO")%>">
<table>
<tr><th>MODEL</th></tr>

<tr><td>
<SELECT size="6" id="color" name="color" OnChange="Buildkey(this.selectedIndex);" multiple>
<%
For i = 0 to UBound(model)
	Response.Write "<option value=" & i & ">" & model(i) & "</option>"
Next
%>
</Select>
</td></tr>

<tr><th>EC_CODE</th></tr>

<tr><td>
<SELECT size="16" id="ec" name="ec" multiple>
<%
arr = Split(model_ec(0),",")
For i = 0 to UBound(arr)
	Response.Write "<option>" & arr(i) & "</option>"
Next
%>
</Select>
</td></tr>

</table>

<input type="submit" value="顯示RunChart" name="B1" style="background-color: #EAF4F4">
<input type="hidden" name="FAB" value="<%=FAB%>">

</form>


<%
If Request("ec") <> "" Then
	Response.Write Request("ec")
End If
%>

</body>
</Html>

<SCRIPT Language="JavaScript"><!--

key=new Array(<%=UBound(model)+1%>);
<%
For i = 0 to UBound(model)
	arr = Split(model_ec(i),",")
	%>
	key[<%=i%>]=new Array(<%=UBound(arr)+1%>);
	<%
	For j = 0 to UBound(arr)
	%>
	key[<%=i%>][<%=j%>]="<%=arr(j)%>";
<%
	Next
Next
%>

function Buildkey(num){
var s= $("#color").val();
var arr = s.toString().split(',');

var ec = ""
$("#ec").empty();
for (i=0;i < arr.length;i++) {
	j = arr[i];
	for (k=0;k < key[j].length;k++) {
		//ec += key[j][k];
		$("#ec").append("<option value='" + key[j][k] + "'>" + key[j][k] + "</option>");
	}
}

document.getElementById("aa").innerHTML = ec;
}

//-->
</Script>


</Body>
</Html>
