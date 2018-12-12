<%
' change to address of your own SMTP server
'email.wickedlyfast.net
strHost = "email.wickedlyfast.net"
  If Request.querystring("Send") <> "" Then
   Set Mail = Server.CreateObject("Persits.MailSender")
   ' enter valid SMTP host
   Mail.Host = strHost

   Mail.From = "sales@greenvillemarine.com" ' From address
   Mail.FromName = "Customer" ' optional
   Mail.AddAddress "sales@greenvillemarine.com"

   ' message subject
   Mail.Subject = "Sales Form"
   ' message body
    Mail.Body = "Sales Form" & VBCRLF
    
	Mail.Body =	Mail.Body   &   "Name:              " & request.form("Name") & VBCRLF
	Mail.Body = Mail.Body   &	"Address:           " & request.form("Address") & VBCRLF
	Mail.Body =	Mail.Body   &	"Email:             " & request.form("Email")& VBCRLF
	Mail.Body =	Mail.Body   &	"Phone:             " & request.form("Phone")& VBCRLF
	Mail.Body =	Mail.Body   &	"Comments:          " & request.form("Comments")& VBCRLF


   
   Mail.Send ' send message
   End If
response.redirect "index.php"
%>

