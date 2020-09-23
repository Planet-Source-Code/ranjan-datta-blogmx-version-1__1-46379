<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/dreamConn.asp" -->
<%
Dim rsBlog__MMColParam
rsBlog__MMColParam = "id"
If (Request.QueryString("id") <> "") Then 
  rsBlog__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim rsBlog
Dim rsBlog_numRows

Set rsBlog = Server.CreateObject("ADODB.Recordset")
rsBlog.ActiveConnection = MM_dreamConn_STRING
rsBlog.Source = "SELECT *, (SELECT COUNT(*) FROM tblComment WHERE tblComment.blogID = joinBlog.BlogID) AS TOTAL_LINKS  FROM joinBlog  WHERE BlogID = " + Replace(rsBlog__MMColParam, "'", "''") + ""
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
rsRss.Source = "SELECT blogImage, blogTitle, blogURL FROM tblRSS"
rsRss.CursorType = 0
rsRss.CursorLocation = 2
rsRss.LockType = 1
rsRss.Open()

rsRss_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">

<head>
<title>dreamlettes - blog link</title>
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

<div id="arch"> 
<h3>Recent Blogs</h3>
<!--#include file="includes/recent.asp" -->
</div>

<div id="nav"> 
<h2>Navigation</h2>
<!--#include file="includes/menu.asp" -->
</div>

<div id="content"> 
<div class="eachBlog">
<h2><%=(rsBlog.Fields.Item("BlogHeadline").Value)%></h2>
<h3><%=(rsBlog.Fields.Item("catName").Value)%> - <%=(rsBlog.Fields.Item("BlogDate").Value)%></h3>
<div class="htmlBlog"><%=RemoveHTML(rsBlog.Fields.Item("BlogHTML").Value)%></div>
<p><a href="permalink.asp?id=<%=(rsBlog.Fields.Item("BlogID").Value)%>">Permalink</a> 
:: <a href="comments.asp?id=<%=(rsBlog.Fields.Item("BlogID").Value)%>" onclick="flevPopupLink(this.href,'winComments','scrollbars=yes,resizable=yes,width=400,height=300',1);return document.MM_returnValue">Comments 
(<%=(rsBlog.Fields.Item("TOTAL_LINKS").Value)%>)</a> :: <a href="#">Top</a></p>
</div>
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
