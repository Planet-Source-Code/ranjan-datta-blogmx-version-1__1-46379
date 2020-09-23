<%
Dim rsRecent__MMColParam
rsRecent__MMColParam = "0"
If (Request("MM_EmptyValue") <> "") Then 
  rsRecent__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsRecent
Dim rsRecent_numRows

Set rsRecent = Server.CreateObject("ADODB.Recordset")
rsRecent.ActiveConnection = MM_dreamConn_STRING
rsRecent.Source = "SELECT BlogHeadline, BlogID FROM tblBlog WHERE BlogIncluded <> " + Replace(rsRecent__MMColParam, "'", "''") + " ORDER BY BlogID DESC"
rsRecent.CursorType = 0
rsRecent.CursorLocation = 2
rsRecent.LockType = 1
rsRecent.Open()

rsRecent_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 5
Repeat1__index = 0
rsRecent_numRows = rsRecent_numRows + Repeat1__numRows
%>
<ul>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsRecent.EOF)) 
%>
<li><a href="permalink.asp?id=<%=(rsRecent.Fields.Item("BlogID").Value)%>"><%=(rsRecent.Fields.Item("BlogHeadline").Value)%></a></li>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsRecent.MoveNext()
Wend
%>
</ul>
<%
rsRecent.Close()
Set rsRecent = Nothing
%>
