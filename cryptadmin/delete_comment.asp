<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_dreamConn_STRING
  MM_editTable = "tblComment"
  MM_editColumn = "commentID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "approve_comments.asp"

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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
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
Dim rsPublish__MMColParam
rsPublish__MMColParam = "1"
If (Request.QueryString("passID") <> "") Then 
  rsPublish__MMColParam = Request.QueryString("passID")
End If
%>

<%
Dim rsPublish
Dim rsPublish_numRows

Set rsPublish = Server.CreateObject("ADODB.Recordset")
rsPublish.ActiveConnection = MM_dreamConn_STRING
rsPublish.Source = "SELECT *  FROM tblComment  WHERE commentID = " + Replace(rsPublish__MMColParam, "'", "''") + ""
rsPublish.CursorType = 0
rsPublish.CursorLocation = 2
rsPublish.LockType = 1
rsPublish.Open()

rsPublish_numRows = 0
%>

<html>
<head>
<title>blogMX :: Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
<label> </label>
<table width="100%"  border="0" cellspacing="1" cellpadding="1">
<tr> 
<td><%=(rsPublish.Fields.Item("commentName").Value)%></td>
<td><%=(rsPublish.Fields.Item("commentEmail").Value)%></td>
<td><input name="hiddenField" type="hidden" value="Yes"> <input name="Submit" type="submit" class="txtBox" value="Delete?"></td>
</tr>
<tr> 
<td colspan="3"><%=(rsPublish.Fields.Item("commentHTML").Value)%></td>
</tr>
</table>
<input type="hidden" name="MM_delete" value="form1">
<input type="hidden" name="MM_recordId" value="<%= rsPublish.Fields.Item("commentID").Value %>">
</form>
</body>
</html>
<%
rsPublish.Close()
Set rsPublish = Nothing
%>

