<%
Dim rsLinks
Dim rsLinks_numRows

Set rsLinks = Server.CreateObject("ADODB.Recordset")
rsLinks.ActiveConnection = MM_dreamConn_STRING
rsLinks.Source = "SELECT * FROM tblLink ORDER BY linkName ASC"
rsLinks.CursorType = 0
rsLinks.CursorLocation = 2
rsLinks.LockType = 1
rsLinks.Open()

rsLinks_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsLinks_numRows = rsLinks_numRows + Repeat1__numRows
%>
<% If Not rsLinks.EOF Or Not rsLinks.BOF Then %>
<ul>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsLinks.EOF)) 
%>
<li><a href="<%=(rsLinks.Fields.Item("linkURL").Value)%>"><%=(rsLinks.Fields.Item("linkName").Value)%></a></li>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsLinks.MoveNext()
Wend
%>
</ul>
<% End If ' end Not rsLinks.EOF Or NOT rsLinks.BOF %>
<%
rsLinks.Close()
Set rsLinks = Nothing
%>
