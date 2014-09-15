<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>邮件发送</title>
</head>

<body>
<%
Function sendEmail("zhengjinyang@stu.pkuschool.edu.cn",Paydisplay)

mailserver="smtp.163.com"  
mailname="somejump@163.com "  
mailpassword="8899174"  
dim msg
CLStr=Chr(13) & Chr(10)
Set msg = Server.CreateObject("JMail.Message")
msg.silent = true
msg.Logging = true
msg.Charset = "gb2312"
msg.MailServerUserName = mailname
msg.MailServerPassword = mailpassword   
msg.From = mailname
msg.FromName = mailname
msg.AddRecipient ("zhengjinyang@stu.pkuschool.edu.cn")          
msg.Subject ="信息反馈"

msg.Body =Paydisplay

msg.Send (mailserver)
msg.close

set msg = nothing
End Function
%>
<%
email=request.Form("T3")
body="报名咨询人: "&request.Form("T1")&" ： <br>"
body=body& "咨询专业："&request.Form("T2")&" <br>"
body=body& "Email 地址："&request.Form("T3")&" <br>"
body=body& "联系电话："&request.Form("T4")&" <br>"
body=body& "详细说明："&request.Form("T5")&" <br>"
Paydisplay=body
'response.Write(sendEmail(email,Paydisplay))'
%>
<%=sendEmail(email,Paydisplay)%>
</body>
</html>