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
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_dreamConn_STRING
  MM_editTable = "tblComment"
  MM_editColumn = "commentID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "reply_comments.asp"
  MM_fieldsStr  = "hiddenField|value"
  MM_columnsStr = "commentInclude|none,Yes,No"

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
<form name="form1" method="POST" action="<%=MM_editAction%>">
<label> </label>
<table width="100%"  border="0" cellspacing="1" cellpadding="1">
<tr> 
<td><%=(rsPublish.Fields.Item("commentName").Value)%></td>
<td><%=(rsPublish.Fields.Item("commentEmail").Value)%></td>
<td><input name="hiddenField" type="hidden" value="Yes"> <input name="Submit" type="submit" class="txtBox" value="Reply & Publish"></td>
</tr>
<tr> 
<td colspan="3"><%=(rsPublish.Fields.Item("commentHTML").Value)%></td>
</tr>
</table>
<input type="hidden" name="MM_update" value="form1">
<input type="hidden" name="MM_recordId" value="<%= rsPublish.Fields.Item("commentID").Value %>">
</form>
</body>
</html>
<%
rsPublish.Close()
Set rsPublish = Nothing
%>

