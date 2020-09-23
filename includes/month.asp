<%
Dim rsDates
Dim rsDates_numRows

Set rsDates = Server.CreateObject("ADODB.Recordset")
rsDates.ActiveConnection = MM_dreamConn_STRING
rsDates.Source = "SELECT DISTINCT Month(BlogDate), Year(BlogDate)  FROM tblBlog  ORDER BY Month(BlogDate) DESC"
rsDates.CursorType = 0
rsDates.CursorLocation = 2
rsDates.LockType = 1
rsDates.Open()

rsDates_numRows = 0
%>

<%
Dim Repeat111__numRows
Dim Repeat111__index

Repeat111__numRows = -1
Repeat111__index = 0
rsDates_numRows = rsDates_numRows + Repeat111__numRows
%>

<ul>

<% 
While ((Repeat111__numRows <> 0) AND (NOT rsDates.EOF)) 
%>
<li><a href="archive.asp?date=<%=(rsDates.Fields.Item("Expr1000").Value)%>/<%=(rsDates.Fields.Item("Expr1001").Value)%>"><%=MonthName((rsDates.Fields.Item("Expr1000").Value))%> - <%=(rsDates.Fields.Item("Expr1001").Value)%></a></li>
<% 
  Repeat111__index=Repeat111__index+1
  Repeat111__numRows=Repeat111__numRows-1
  rsDates.MoveNext()
Wend
%>
</ul>

<%
rsDates.Close()
Set rsDates = Nothing
%>
