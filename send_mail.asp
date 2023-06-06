	Set mail = Server.CreateObject( "CDONTS.NewMail" )
	mail.To = manager_mail
	mail.From = Request("Emp_Mail")
	mail.Cc = "jimmyzmwu@auo.com;" & Request("Emp_Mail")
	mail.Subject = "製造部公告，請主管確認與簽核"
	mail.BodyFormat = 0
	mail.MailFormat = 0
	mail.Body = "主旨：" & Request("Subject") & "<br>內容：" & Request("Words") & "<br><a href=http://10.88.18.23/Contents/製造部公告/detail.asp?TitleID=" & TitleID & ">按此簽核</a>"
	mail.Send
