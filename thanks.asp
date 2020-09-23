<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Send Form Values From Previous Page to Webmaster - CDOMail_112
Dim usxCDO
Set usxCDO = Server.CreateObject("CDONTS.NewMail")
usxCDO.From = "ranjan@dreamlettes.net"
usxCDO.To = "admin@dreamlettes.net"
usxCDO.Subject = "New Comment on Blog"
usxCDO.Body = Chr(13) & Chr(10) &_
"There is a new comment on your blog awaiting approval. "
usxCDO.Send
Set usxCDO = Nothing 
'Response.Redirect "a"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Thanks</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<h2>Thanks</h2>
<p>Thank you for your post. It will appear on the comments list on approval</p>
<p><a href="javascript:window.close()">Click here</a> to close</p>
</body>
</html>
