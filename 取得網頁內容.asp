ServerURL="http://10.88.38.99:5002/?s=" + S
Set Mail1 = Server.CreateObject("CDO.Message")
Mail1.CreateMHTMLBody ServerURL
reply=Mail1.HTMLBody
Set Mail1 = Nothing

ServerURL="http://c5aweb01/MFG/Personal/Personal_Info.asp?WorkID=" & Session("WorkID")
Set Mail1 = Server.CreateObject("CDO.Message") 
Mail1.CreateMHTMLBody ServerURL,31 
AA=Mail1.HTMLBody '結果回傳到AA裡 
Set Mail1 = Nothing
message_count = Split(AA,",")
message_count = Array(0,0)
