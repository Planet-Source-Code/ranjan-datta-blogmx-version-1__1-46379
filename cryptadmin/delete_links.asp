<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="default.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<!--#include file="../Connections/dreamConn.asp" -->
<%

if(Request("delBox") <> "") then cmdDelLinks__delLinkIDs = Request("delBox")

%>
<%

set cmdDelLinks = Server.CreateObject("ADODB.Command")
cmdDelLinks.ActiveConnection = MM_dreamConn_STRING
cmdDelLinks.CommandText = "DELETE FROM tblLink  WHERE linkID IN (" + Replace(cmdDelLinks__delLinkIDs, "'", "''") + ") "
cmdDelLinks.CommandType = 1
cmdDelLinks.CommandTimeout = 0
cmdDelLinks.Prepared = true
cmdDelLinks.Execute()

%>
<%
Dim rsLinks
Dim rsLinks_numRows

Set rsLinks = Server.CreateObject("ADODB.Recordset")
rsLinks.ActiveConnection = MM_dreamConn_STRING
rsLinks.Source = "SELECT * FROM tblLink"
rsLinks.CursorType = 0
rsLinks.CursorLocation = 2
rsLinks.LockType = 1
rsLinks.Open()

rsLinks_numRows = 0
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
	   "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>blogMX :: Administration</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="robots" content="NOINDEX, NOFOLLOW">
</head>

<body>
<table width="100%"  border="0" cellspacing="2" cellpadding="3">
<tr>
<td>&nbsp;</td>
<td>&nbsp;</td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td><h1>Delete Links</h1></td>
</tr>
<tr>
<td valign="top" id="nav"> 
<!--#include file="menu.asp" -->
</td>
<td valign="top">Selected Links were deleted. <a href="main.asp">Back to Main</a> 
Page?</td>
</tr>
</table>
</body>
</html>
<%
rsLinks.Close()
Set rsLinks = Nothing
%>

