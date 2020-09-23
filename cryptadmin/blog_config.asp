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
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_dreamConn_STRING
  MM_editTable = "tblRSS"
  MM_editColumn = "rssID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "main.asp"
  MM_fieldsStr  = "txtBlog|value|txtDesc|value|txtURL|value|txtImg|value|txtCopy|value|txtMail|value|txtCal|value|txtMon|value|txtCat|value|txtSea|value"
  MM_columnsStr = "blogTitle|',none,''|blogDesc|',none,''|blogURL|',none,''|blogImage|',none,''|blogCopyRight|',none,''|blogEmail|',none,''|incCalendar|none,Yes,No|incMonth|none,Yes,No|incCat|none,Yes,No|incSearch|none,Yes,No"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsBlogConfig
Dim rsBlogConfig_numRows

Set rsBlogConfig = Server.CreateObject("ADODB.Recordset")
rsBlogConfig.ActiveConnection = MM_dreamConn_STRING
rsBlogConfig.Source = "SELECT * FROM tblRSS"
rsBlogConfig.CursorType = 0
rsBlogConfig.CursorLocation = 2
rsBlogConfig.LockType = 1
rsBlogConfig.Open()

rsBlogConfig_numRows = 0
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
<td><h1>Configure Blog</h1></td>
</tr>
<tr>
<td valign="top" id="nav"> 
<!--#include file="menu.asp" -->
</td>
<td valign="top"> 
<form name="form1" method="POST" action="<%=MM_editAction%>">
<table width="100%"  border="0">
<tr> 
<td>Blog Title</td>
<td><input name="txtBlog" type="text" id="txtBlog" value="<%=(rsBlogConfig.Fields.Item("blogTitle").Value)%>"></td>
</tr>
<tr> 
<td>Blog Description</td>
<td><textarea name="txtDesc" cols="30" rows="5" id="txtDesc"><%=(rsBlogConfig.Fields.Item("blogDesc").Value)%></textarea></td>
</tr>
<tr> 
<td>Blog URL</td>
<td><input name="txtURL" type="text" id="txtURL" value="<%=(rsBlogConfig.Fields.Item("blogURL").Value)%>"></td>
</tr>
<tr> 
<td>Blog Image</td>
<td><input name="txtImg" type="text" id="txtImg" value="<%=(rsBlogConfig.Fields.Item("blogImage").Value)%>"></td>
</tr>
<tr> 
<td>Blog Copyright</td>
<td><input name="txtCopy" type="text" id="txtCopy" value="<%=(rsBlogConfig.Fields.Item("blogCopyRight").Value)%>"></td>
</tr>
<tr> 
<td>Blog Email</td>
<td><input name="txtMail" type="text" id="txtMail" value="<%=(rsBlogConfig.Fields.Item("blogEmail").Value)%>"></td>
</tr>
<tr> 
<td colspan="2"> <input <%If rsBlogConfig.Fields.Item("incCalendar").Value = -1 Then Response.Write("checked") %> name="txtCal" type="checkbox" id="txtCal" value="-1">
Include Calendar <input <%If rsBlogConfig.Fields.Item("incMonth").Value = -1 Then Response.Write("checked") %> name="txtMon" type="checkbox" id="txtMon" value="-1">
Include Months <input <%If rsBlogConfig.Fields.Item("incCat").Value = -1 Then Response.Write("checked") %> name="txtCat" type="checkbox" id="txtCat" value="-1">
Include Category <input <%If rsBlogConfig.Fields.Item("incSearch").Value = -1 Then Response.Write("checked") %> name="txtSea" type="checkbox" id="txtSea" value="-1">
Include Search</td>
</tr>
<tr> 
<td>&nbsp;</td>
<td> <input type="submit" name="Submit" value="Submit"></td>
</tr>
</table>
<input type="hidden" name="MM_update" value="form1">
<input type="hidden" name="MM_recordId" value="<%= rsBlogConfig.Fields.Item("rssID").Value %>">
</form></td>
</tr>
</table>
</body>
</html>
<%
rsBlogConfig.Close()
Set rsBlogConfig = Nothing
%>
