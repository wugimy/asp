ServerURL="http://10.88.38.99:5002/?s=" + S
Set Mail1 = Server.CreateObject("CDO.Message")
Mail1.CreateMHTMLBody ServerURL
reply=Mail1.HTMLBody
Set Mail1 = Nothing
