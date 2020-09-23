<%
Dim rsCat
Dim rsCat_numRows

Set rsCat = Server.CreateObject("ADODB.Recordset")
rsCat.ActiveConnection = MM_dreamConn_STRING
rsCat.Source = "SELECT * FROM tblCategory ORDER BY catName ASC"
rsCat.CursorType = 0
rsCat.CursorLocation = 2
rsCat.LockType = 1
rsCat.Open()

rsCat_numRows = 0
%>
<form id="form1" method="get" action="archive.asp?cat=<%=(rsCat.Fields.Item("catID").Value)%>">
<fieldset>
<label for="cat"></label>
<select name="cat" id="cat" onchange="this.form.submit()">
<option selected="selected" value="">Select Category</option>
<%
While (NOT rsCat.EOF)
%>
<option value="<%=(rsCat.Fields.Item("catID").Value)%>"><%=(rsCat.Fields.Item("catName").Value)%></option>
<%
  rsCat.MoveNext()
Wend
If (rsCat.CursorType > 0) Then
  rsCat.MoveFirst
Else
  rsCat.Requery
End If
%>
</select>
</fieldset>
</form>

<%
rsCat.Close()
Set rsCat = Nothing
%>
