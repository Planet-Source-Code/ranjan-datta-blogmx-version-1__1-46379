<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/dreamConn.asp" -->
<%
Response.Buffer = true
if Request("styleSelect") <> "" then
Response.Cookies("style") = Request("styleSelect")
Response.Cookies("style").Expires = _
FormatDateTime(DateAdd("d", 60, Now()))
On Error Resume Next
else
Response.Cookies("style").Expires = _
FormatDateTime(DateAdd("d", -1, Now()))
end if
Function ExtractFileName_InStrRev(ByVal strPath)
Dim strTemp
strTemp = Mid(strPath, InStrRev(strPath, "/") + 1)
ExtractFileName_InStrRev = strTemp
End Function
dim style_file
style_file=ExtractFileName_InStrRev(request.servervariables("url"))
%>


<%
Dim rsBlog
Dim rsBlog_numRows

Set rsBlog = Server.CreateObject("ADODB.Recordset")
rsBlog.ActiveConnection = MM_dreamConn_STRING
rsBlog.Source = "SELECT *, (SELECT COUNT(*) FROM tblComment WHERE tblComment.blogID = tblBlog.BlogID AND tblComment.commentInclude <> 0) AS TOTAL_LINKS  FROM joinBlog  WHERE BlogIncluded <> 0  ORDER BY BlogID DESC"
rsBlog.CursorType = 0
rsBlog.CursorLocation = 2
rsBlog.LockType = 1
rsBlog.Open()

rsBlog_numRows = 0
%>
<%
Dim rsRss
Dim rsRss_numRows

Set rsRss = Server.CreateObject("ADODB.Recordset")
rsRss.ActiveConnection = MM_dreamConn_STRING
rsRss.Source = "SELECT * FROM tblRSS"
rsRss.CursorType = 0
rsRss.CursorLocation = 2
rsRss.LockType = 1
rsRss.Open()

rsRss_numRows = 0
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
<title>blog MX</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<script type="text/javascript" src="includes/allscripts.js"></script>
<% if Request.Cookies("style") <> "" then %>
<style type="text/css">
<!--
@import url("styles/<% = Request.Cookies("style") %>.css");
-->
</style>
<% else %>
<style type="text/css">
<!--
@import url("styles/plainblue.css");
-->
</style>
<% end if %>
</head>

<body>
<div id="header"> 
<h1><a href="<%=(rsRss.Fields.Item("blogURL").Value)%>"><img src="<%=(rsRss.Fields.Item("blogimage").value)%>" alt="<%=(rsRss.Fields.Item("blogTitle").Value)%>" /><%=(rsRss.Fields.Item("blogTitle").Value)%></a></h1>
</div>

<div id="arch" class="btn"> 
<% If  (rsRss.Fields.Item("incCalendar").Value) <> 0 Then %>
<h3>Blog Calendar</h3>
<!--#include file="includes/calendar.asp" -->
<% End If %>
<% If (rsRss.Fields.Item("incMonth").Value) <> 0 Then %>
<h3>Blogs by Month</h3>
<!--#include file="includes/month.asp" -->
<% End If %>
<% If (rsRss.Fields.Item("incCat").Value) <> 0 Then %>
<h3>Blogs by Category</h3>
<!--#include file="includes/category.asp" -->
<% End If %>
<% If (rsRss.Fields.Item("incSearch").Value) <> 0 Then %>
<h3>Blog Search</h3>
<!--#include file="includes/search.asp" -->
<% End If %>
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
<h2>Select Style</h2>
<form id="styleFrm" method="post" action="<%=style_file%>">
<fieldset>
<label for="styleSelect"></label>
<select name="styleSelect" id="styleSelect" onChange="this.form.submit()">
<option selected="selected">Select Style 
<option value="basic">basic 
<option value="chocolate">chocolate 
<option value="plainblue">plainblue
</select>
</fieldset>
</form>
</div>

<div id="content">
<% 
While ((Repeat2__numRows <> 0) AND (NOT rsBlog.EOF)) 
%>
<div class="eachBlog">
<h2><%=(rsBlog.Fields.Item("BlogHeadline").Value)%></h2>
<h3><%=(rsBlog.Fields.Item("catName").Value)%> - <%=(rsBlog.Fields.Item("BlogDate").Value)%></h3>
<div class="htmlBlog"><%=(RemoveHTML(rsBlog.Fields.Item("BlogHTML").Value))%></div>
<p><a href="permalink.asp?id=<%=(rsBlog.Fields.Item("BlogID").Value)%>">Permalink</a> 
:: <a href="comments.asp?id=<%=(rsBlog.Fields.Item("BlogID").Value)%>" onClick="flevPopupLink(this.href,'winComments','scrollbars=yes,resizable=yes,width=400,height=300',1);return document.MM_returnValue">Comments 
(<%=(rsBlog.Fields.Item("TOTAL_LINKS").Value)%>)</a> :: <a href="#">Top</a></p>
</div>
<% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsBlog.MoveNext()
Wend
%>
</div>

<div id="footer"> 
<p>powered by blogMX - <a href="http://www.dreamlettes.net">http://www.dreamlettes.net 
</a> </p>
</div>
</body>
</html>
<%
rsBlog.Close()
Set rsBlog = Nothing
%>
<%
rsRss.Close()
Set rsRss = Nothing
%>
