<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/dreamConn.asp" -->
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
<%
Dim rsComments__MMColParam
rsComments__MMColParam = "0"
If (Request("MM_EmptyValue") <> "") Then 
  rsComments__MMColParam = Request("MM_EmptyValue")
End If
%>



<%
Dim rsComments
Dim rsComments_numRows

Set rsComments = Server.CreateObject("ADODB.Recordset")
rsComments.ActiveConnection = MM_dreamConn_STRING
rsComments.Source = "SELECT *  FROM tblComment  WHERE commentInclude = " + Replace(rsComments__MMColParam, "'", "''") + "  ORDER BY commentDate ASC"
rsComments.CursorType = 0
rsComments.CursorLocation = 2
rsComments.LockType = 1
rsComments.Open()

rsComments_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsComments_numRows = rsComments_numRows + Repeat1__numRows
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
<td><h1>Approve Comments</h1></td>
</tr>
<tr>
<td valign="top" id="nav"> <p>
<!--#include file="menu.asp" -->
</p>
</td>
<td>

<% If Not rsComments.EOF Or Not rsComments.BOF Then %>
<table width="100%"  border="0" cellspacing="1" cellpadding="1">

<% 
While ((Repeat1__numRows <> 0) AND (NOT rsComments.EOF)) 
%>
<tr>
<td><a href="<%=(rsComments.Fields.Item("commentURL").Value)%>"><%=(rsComments.Fields.Item("commentName").Value)%></a></td>
<td><%=(rsComments.Fields.Item("commentEmail").Value)%></td>
<td><a href="confirm_publish.asp?passID=<%=(rsComments.Fields.Item("commentID").Value)%>">Approve?</a> / <a href="delete_comment.asp?passID=<%=(rsComments.Fields.Item("commentID").Value)%>">Delete?</a></td>
</tr>
<tr>
<td colspan="3"><%=(rsComments.Fields.Item("commentHTML").Value)%></td>
</tr>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsComments.MoveNext()
Wend
%>


</table>
<% End If ' end Not rsComments.EOF Or NOT rsComments.BOF %>
</td>
</tr>
</table>
</body>
</html>
<%
rsComments.Close()
Set rsComments = Nothing
%>





