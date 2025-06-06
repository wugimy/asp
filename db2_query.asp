set cnn=server.createobject("adodb.connection")
If FAB = "A" Then
	cnn.open "DSN=L8ADAHS1;driver={IBM DB2 ODBC Driver};DBALIAS=L8ADAHS1; uid=auordary;pwd=Auo$1"
ElseIf FAB = "B" Then
	cnn.open "DSN=ADAHS1B;driver={IBM DB2 ODBC Driver};DBALIAS=ADAHS1B; uid=auordary;pwd=Adt$1"
End If
