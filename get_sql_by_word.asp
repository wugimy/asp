Sub get_sql_by_word(WORD)
	SQL = "select QUERY_SQL,SUBSTR(QUERY_SQL,LOCATE('{',QUERY_SQL),LOCATE('}',QUERY_SQL)-LOCATE('{',QUERY_SQL)+1) as WORD_A,SUBSTR(QUERY_SQL,LOCATE('{',QUERY_SQL,LOCATE('}',QUERY_SQL)+1),LOCATE('}',QUERY_SQL,LOCATE('}',QUERY_SQL)+1)-LOCATE('{',QUERY_SQL,LOCATE('}',QUERY_SQL)+1)+1) as WORD_B from word_to_sql where WORD='" & WORD & "'"
	SQL = "select A.*,B.QUERY_SQL as SQL_A from (" & SQL & ") A left outer join word_to_sql B on A.WORD_A=CONCAT('{',B.WORD,'}')"
	SQL = "select A.*,B.QUERY_SQL as SQL_B from (" & SQL & ") A left outer join word_to_sql B on A.WORD_B=CONCAT('{',B.WORD,'}')"
	SQL = "select REPLACE(REPLACE(QUERY_SQL,WORD_A,CONCAT('(',SQL_A,')')),WORD_B,CONCAT('(',SQL_B,')')) from (" & SQL & ")A"

    set rs=server.createobject("adodb.Recordset")
    rs.Open SQL,cnn
	
	If Not rs.EOF Then
		NEW_SQL = rs(0)
	End If
	
    rs.Close
    set rs = nothing
End Sub
