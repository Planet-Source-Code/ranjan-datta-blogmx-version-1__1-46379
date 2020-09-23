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

if(Request("delBox") <> "") then cmdDelBlog__delBlogIDs = Request("delBox")

%>
<%

set cmdDelBlog = Server.CreateObject("ADODB.Command")
cmdDelBlog.ActiveConnection = MM_dreamConn_STRING
cmdDelBlog.CommandText = "DELETE FROM tblBlog  WHERE BlogID IN (" + Replace(cmdDelBlog__delBlogIDs, "'", "''") + ") "
cmdDelBlog.CommandType = 1
cmdDelBlog.CommandTimeout = 0
cmdDelBlog.Prepared = true
cmdDelBlog.Execute()

%>
<%
Dim rsBlog
Dim rsBlog_numRows

Set rsBlog = Server.CreateObject("ADODB.Recordset")
rsBlog.ActiveConnection = MM_dreamConn_STRING
rsBlog.Source = "SELECT BlogDate, BlogHeadline, BlogID FROM tblBlog ORDER BY BlogID DESC"
rsBlog.CursorType = 0
rsBlog.CursorLocation = 2
rsBlog.LockType = 1
rsBlog.Open()

rsBlog_numRows = 0
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
<td><h1>Delete Blogs</h1></td>
</tr>
<tr>
<td valign="top" id="nav"> 
<!--#include file="menu.asp" -->
</td>
<td valign="top">Selected Blogs were deleted. <a href="main.asp">Back to Main</a> 
Page?</td>
</tr>
</table>
</body>
</html>
<%
rsBlog.Close()
Set rsBlog = Nothing
%>

