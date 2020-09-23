<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/dreamConn.asp" -->
<%
Dim rsRss
Dim rsRss_numRows

Set rsRss = Server.CreateObject("ADODB.Recordset")
rsRss.ActiveConnection = MM_dreamConn_STRING
rsRss.Source = "SELECT blogImage, blogTitle, blogURL FROM tblRSS"
rsRss.CursorType = 0
rsRss.CursorLocation = 2
rsRss.LockType = 1
rsRss.Open()

rsRss_numRows = 0
%>
<%
Dim extMonth
Dim extYear
Dim extDay
If Request.QueryString("date") <> "" Then
strLoc= InStr(Request.QueryString("date"), "/")
extMonth=Left(Request.QueryString("date"), strLoc-1)
extYear=Right(Request.QueryString("date"),4)
End If
If Request.QueryString("month") <> "" Then
extMonth=Request.QueryString("month")
extDay = Request.QueryString("day")
extYear=Request.QueryString("year")
End IF
%>
<%
'Function PCase(strInput)
'    Dim iPosition ' Our current position in the string (First character = 1)
'    Dim iSpace ' The position of the next space after our iPosition
'    Dim strOutput ' Our temporary string used to build the function's output

 '   iPosition = 1

  '  Do While InStr(iPosition, strInput, " ", 1) <> 0
 '           iSpace = InStr(iPosition, strInput, " ", 1)
 '           strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
 '           strOutput = strOutput & LCase(Mid(strInput, iPosition + 1, iSpace - iPosition))
 '           iPosition = iSpace + 1
 '   Loop

 '   strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
 '   strOutput = strOutput & LCase(Mid(strInput, iPosition + 1))

 '   PCase = strOutput
'End Function
%> 
<%
Dim rsBlog__varSearch
rsBlog__varSearch = "0"
If (Request("txtKeyWord")  <> "") Then 
	rsBlog__varSearch = Request("txtKeyWord") 
	searchArray = Split(Request("txtKeyWord")," ")
	startSQL = "SELECT *, (SELECT COUNT(*) FROM tblComment WHERE tblComment.blogID = tblBlog.BlogID AND tblComment.commentInclude <> 0) AS TOTAL_LINKS  FROM joinBlog  WHERE BlogIncluded <> 0 AND BlogHTML LIKE '%"
	endSQL = "%'  ORDER BY BlogID DESC"
	for i = 0 to Ubound(SearchArray)
		if i > 0 then
		'Builds the sql statement using the CompType to substitute AND/OR
		searchSQL = searchSQL + searchArray(i) + "%' OR BlogHeadline LIKE '%" + searchArray(i)
		else
		'Ends the sql statement if there is only one word
		searchSQL = Replace(rsBlog__varSearch, "'", "''") + "%' OR BlogHeadline LIKE '%" + Replace(rsBlog__varSearch, "'", "''")
		end if
	next
	totalSQL = startSQL + searchSQL + endSQL
End If
%>

<%
Dim rsBlog__MMMonth
rsBlog__MMMonth = "Month(date())"
If (extMonth <> "") Then 
  rsBlog__MMMonth = extMonth
End If
%>
<%
Dim rsBlog__MMYear
rsBlog__MMYear = "Year(date())"
If (extYear <> "") Then 
  rsBlog__MMYear = extYear
End If
%>
<%
Dim rsBlog__MMCat
rsBlog__MMCat = "NULL"
If (Request.QueryString("cat") <> "") Then 
  rsBlog__MMCat = Request.QueryString("cat")
End If
%>
<%
mySQL = "SELECT *, (SELECT COUNT(*) FROM tblComment WHERE tblComment.blogID = tblBlog.BlogID AND tblComment.commentInclude <> 0) AS TOTAL_LINKS  FROM joinBlog  WHERE BlogIncluded <> 0 AND Month(BlogDate) = " + Replace(rsBlog__MMMonth, "'", "''") + " AND Year(BlogDate) = " + Replace(rsBlog__MMYear, "'", "''") + " ORDER BY BlogID DESC"
If Request.QueryString("date") <> "" Then
mySQL = "SELECT *, (SELECT COUNT(*) FROM tblComment WHERE tblComment.blogID = tblBlog.BlogID AND tblComment.commentInclude <> 0) AS TOTAL_LINKS  FROM joinBlog  WHERE BlogIncluded <> 0 AND Month(BlogDate) = " + Replace(rsBlog__MMMonth, "'", "''") + " AND Year(BlogDate) = " + Replace(rsBlog__MMYear, "'", "''") + " ORDER BY BlogID DESC"
End If
If Request.QueryString("cat") <> "" Then
mySQL = "SELECT *, (SELECT COUNT(*) FROM tblComment WHERE tblComment.blogID = tblBlog.BlogID AND tblComment.commentInclude <> 0) AS TOTAL_LINKS  FROM joinBlog  WHERE BlogIncluded <> 0 AND " + Replace(rsBlog__MMCat, "'", "''") + " = tblBlog_CatID  ORDER BY BlogID DESC"
End If
If Request.QueryString("month") <> "" Then
mySQL = "SELECT *, (SELECT COUNT(*) FROM tblComment WHERE tblComment.blogID = tblBlog.BlogID AND tblComment.commentInclude <> 0) AS TOTAL_LINKS  FROM joinBlog  WHERE BlogIncluded <> 0 AND Month(BlogDate) = " + Replace(rsBlog__MMMonth, "'", "''") + "AND Day(BlogDate) =" + extDay + " AND Year(BlogDate) = " + Replace(rsBlog__MMYear, "'", "''") + " ORDER BY BlogID DESC"
End If
If Request.Form("txtKeyWord") <> "" Then
'mySQL = "SELECT *, (SELECT COUNT(*) FROM tblComment WHERE tblComment.blogID = tblBlog.BlogID AND tblComment.commentInclude <> 0) AS TOTAL_LINKS  FROM joinBlog  WHERE BlogIncluded <> 0 AND BlogHTML LIKE '%" + Replace(rsBlog__varSearch, "'", "''") + "%' OR BlogHeadline LIKE '%" + Replace(rsBlog__varSearch, "'", "''") + "%'  ORDER BY BlogID DESC"
mySQL = totalSQL
End If
%>
<%
Dim rsBlog
Dim rsBlog_numRows

Set rsBlog = Server.CreateObject("ADODB.Recordset")
rsBlog.ActiveConnection = MM_dreamConn_STRING
rsBlog.Source = mySQL
rsBlog.CursorType = 0
rsBlog.CursorLocation = 2
rsBlog.LockType = 1
rsBlog.Open()

rsBlog_numRows = 0
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = 10
Repeat2__index = 0
rsBlog_numRows = rsBlog_numRows + Repeat2__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">

<head>
<title>dreamlettes - blog archives</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<script  type="text/JavaScript" src="includes/allscripts.js"></script>
<% if Request.Cookies("style") <> "" then %>
<style type="text/css">
<!--
@import url("styles/<% = Request.Cookies("style") %>.css");
-->
</style>
<% else %>
<style type="text/css">
<!--
@import url("styles/basic.css");
-->
</style>
<% end if %>
</head>

<body>
<div id="header"> 
<h1><a href="<%=(rsRss.Fields.Item("blogURL").Value)%>"><img src="<%=(rsRss.Fields.Item("blogimage").value)%>" alt="<%=(rsRss.Fields.Item("blogTitle").Value)%>" /><%=(rsRss.Fields.Item("blogTitle").Value)%></a></h1>
</div>
<div id="arch" class="btn"> 
<h3>Blog Calendar</h3>
<!--#include file="includes/calendar.asp" -->
<h3>Blogs by Month</h3>
<!--#include file="includes/month.asp" -->
<h3>Blogs by Category</h3>
<!--#include file="includes/category.asp" -->
<h3>Blog Search</h3>
<!--#include file="includes/search.asp" -->
</div>
<div id="nav"> 
<h2>Navigation</h2>
<!--#include file="includes/menu.asp" -->
<h2>Blog Links</h2>
<!--#include file="includes/links.asp" -->
<h2>RSS Feed</h2>
<ul>
<li><a href="rss.asp">RSS 2.0</a></li>
</ul>
</div>

<div id="content"> 
<% If rsBlog.EOF And rsBlog.BOF Then %>
<h2>No Records for search criteria</h2>
<% End If %>
<% If Not rsBlog.EOF Or Not rsBlog.BOF Then %>
<% 
While ((Repeat2__numRows <> 0) AND (NOT rsBlog.EOF)) 
%>
<div class="eachBlog">
<h2><%=(rsBlog.Fields.Item("BlogHeadline").Value)%></h2>
<h3><%=(rsBlog.Fields.Item("catName").Value)%> - <%=(rsBlog.Fields.Item("BlogDate").Value)%></h3>
<div class="htmlBlog"><%=(RemoveHTML(rsBlog.Fields.Item("BlogHTML").Value))%></div><p><a href="permalink.asp?id=<%=(rsBlog.Fields.Item("BlogID").Value)%>">Permalink</a> 
:: <a href="comments.asp?id=<%=(rsBlog.Fields.Item("BlogID").Value)%>" onClick="flevPopupLink(this.href,'winComments','scrollbars=yes,resizable=yes,width=400,height=300',1);return document.MM_returnValue">Comments 
(<%=(rsBlog.Fields.Item("TOTAL_LINKS").Value)%>)</a> :: <a href="#">Top</a></p>
</div>
<% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsBlog.MoveNext()
Wend
%>
<% End If ' end Not rsBlog.EOF Or NOT rsBlog.BOF %>
</div>

<div id="footer"> 
<p>powered by blogMX - <a href="http://www.dreamlettes.net">http://www.dreamlettes.net 
</a> </p>
</div>
</body>
</html>
<%
rsRss.Close()
Set rsRss = Nothing
%>
<%
rsBlog.Close()
Set rsBlog = Nothing
%>

