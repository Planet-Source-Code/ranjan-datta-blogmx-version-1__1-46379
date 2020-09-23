<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="Connections/dreamConn.asp" -->
<%
If cstr(Request.Form("txtName"))<>"" Then
If Request.form("remember") ="1" Then
Response.Cookies("ckName") = Request.Form("txtName")
Response.Cookies("ckURL") = Request.Form("txtURL")
Response.Cookies("ckEmail") = Request.Form("txtEmail")
Response.Cookies("ckRemember") = "1"
Response.Cookies("ckName").Expires = Date + 30
Response.Cookies("ckURL").Expires = Date + 30
Response.Cookies("ckEmail").Expires = Date + 30
Response.Cookies("ckRemember").expires = Date + 30
Else
Response.Cookies("ckName") = ""
Response.Cookies("ckURL") = ""
Response.Cookies("ckEmail") = ""
Response.Cookies("ckRemember") = ""
End If
End If
%>
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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_dreamConn_STRING
  MM_editTable = "tblComment"
  MM_editRedirectUrl = "thanks.asp"
  MM_fieldsStr  = "txtName|value|txtURL|value|txtEmail|value|textarea|value|hiddenField|value"
  MM_columnsStr = "commentName|',none,''|commentURL|',none,''|commentEmail|',none,''|commentHTML|',none,''|blogID|none,none,NULL"

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
Dim rsComments__MMColParam
rsComments__MMColParam = "0"
If (Request("MM_EmptyValue") <> "") Then 
  rsComments__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsComments__MMColParam2
rsComments__MMColParam2 = "0"
If (Request.QueryString("id")   <> "") Then 
  rsComments__MMColParam2 = Request.QueryString("id")  
End If
%>
<%
Dim rsComments
Dim rsComments_numRows

Set rsComments = Server.CreateObject("ADODB.Recordset")
rsComments.ActiveConnection = MM_dreamConn_STRING
rsComments.Source = "SELECT *  FROM tblComment  WHERE commentInclude <> " + Replace(rsComments__MMColParam, "'", "''") + " AND blogID = " + Replace(rsComments__MMColParam2, "'", "''") + "  ORDER BY commentID DESC"
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


<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<title>blogMX :: Comments</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<script  type="text/javascript" src="includes/allscripts.js">
</script>
</head>

<body>
<h2>Comments</h2>

<% 
While ((Repeat1__numRows <> 0) AND (NOT rsComments.EOF)) 
%>
<p><%=(rsComments.Fields.Item("commentName").Value)%> says...</p>
<div><%=RemoveHTML(rsComments.Fields.Item("commentHTML").Value)%> </div>
<% If  (rsComments.Fields.Item("commentURL").Value) <> "" Then %>
<p>Posted by <a href="<%=(rsComments.Fields.Item("commentURL").Value)%>" onclick="flevPopupLink(this.href,'guestWin','toolbar=yes,location=yes,status=yes,menubar=yes,scrollbars=yes,resizable=yes',0);return document.MM_returnValue"><%=(rsComments.Fields.Item("commentName").Value)%></a> 
<% Else %>
<p>Posted by <%=(rsComments.Fields.Item("commentName").Value)%> 
<% End If %>
&nbsp;on <%=(rsComments.Fields.Item("commentDate").Value)%></p>
<hr>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsComments.MoveNext()
Wend
%>
<form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
<table width="90%"  border="0" cellspacing="2" cellpadding="3">
<tr> 
<td align="right" valign="top">Name</td>
<td> <input name="txtName" type="text" id="txtName" value="<%= Request.Cookies("ckName") %>" /> 
</td>
</tr>
<tr> 
<td align="right" valign="top">URL</td>
<td> <input name="txtURL" type="text" id="txtURL" value="<%= Request.Cookies("ckURL") %>" /></td>
</tr>
<tr> 
<td align="right" valign="top">Email</td>
<td> <input name="txtEmail" type="text" id="txtEmail" value="<%= Request.Cookies("ckEmail") %>" />
<br />
Email address is not pulished</td>
</tr>
<tr> 
<td align="right" valign="top">Remember Me</td>
<td><input name="remember" type="checkbox" id="remember" value="1" checked="checked" <%If (Request.Cookies("ckRemember") = "1") Then Response.Write("CHECKED") : Response.Write("")%> /></td>
</tr>
<tr> 
<td align="right" valign="top">Comments<br />
</td>
<td> <textarea name="textarea" rows="10"></textarea></td>
</tr>
<tr> 
<td valign="top"> <input name="hiddenField" type="hidden" value="<%= Request.Querystring("id") %>" /></td>
<td> <input name="Submit" type="submit" onclick="YY_checkform('form1','txtName','#q','0','Name is required*','textarea','1','1','Duh? Comment Something*');return document.MM_returnValue" value="Comment" /></td>
</tr>
</table>
<input type="hidden" name="MM_insert" value="form1" />
</form>
</body>
</html>
<%
rsComments.Close()
Set rsComments = Nothing
%>

