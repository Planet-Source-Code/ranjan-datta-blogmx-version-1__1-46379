<%@LANGUAGE="VBSCRIPT"%>
<%
Dim MM_dreamConn_STRING
MM_dreamConn_STRING = "DSN=dsnBlogMX"
Dim objRS
Dim objRS_numRows
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.ActiveConnection = MM_dreamConn_STRING
objRS.Source = "SELECT BlogDate, BlogHeadline, BlogHTML, BlogID FROM joinBlog ORDER BY BlogID DESC"
objRS.CursorType = 0
objRS.CursorLocation = 2
objRS.LockType = 1
objRS.Open()
objRS_numRows = 0
Dim rsSiteDetails
Dim rsSiteDetails_numRows
Set rsSiteDetails = Server.CreateObject("ADODB.Recordset")
rsSiteDetails.ActiveConnection = MM_dreamConn_STRING
rsSiteDetails.Source = "SELECT * FROM tblRSS"
rsSiteDetails.CursorType = 0
rsSiteDetails.CursorLocation = 2
rsSiteDetails.LockType = 1
rsSiteDetails.Open()
rsSiteDetails_numRows = 0
Dim Repeat1__numRows
Dim Repeat1__index
Repeat1__numRows = 10
Repeat1__index = 0
objRS_numRows = objRS_numRows + Repeat1__numRows
  'Output the XML
  Response.ContentType = "text/xml"
  Response.Write("<?xml version=""1.0"" encoding=""utf-8"" ?> ")
  Response.Write("<rss version=""2.0"">")
  Response.Write("<channel>")
  
  Response.Write("<title>" & (rsSiteDetails.Fields.Item("blogTitle").Value) & "</title>") 
  Response.Write("<description>") 
  Response.Write("<![CDATA[" & (rsSiteDetails.Fields.Item("blogDesc").Value) & "]]>")
  Response.Write("</description>") 
  Response.Write("<link>" & (rsSiteDetails.Fields.Item("blogURL").Value) & "</link> ") 
  Response.Write("<language>en</language>") 
  Response.Write("<copyright>") 
  Response.Write("<![CDATA[" & (rsSiteDetails.Fields.Item("blogCopyRight").Value) & "]]>")
  Response.Write("</copyright>") 
  Response.Write("<image>")
  Response.Write("<title>" & (rsSiteDetails.Fields.Item("blogTitle").Value) & "</title>")
  Response.Write("<url>" & (rsSiteDetails.Fields.Item("blogImage").Value) & "</url>") 
  Response.Write("<link>" & (rsSiteDetails.Fields.Item("blogURL").Value) & "</link>")
  Response.Write("</image>")
While ((Repeat1__numRows <> 0) AND (NOT objRS.EOF)) 
  Response.Write("<item>")
  Response.Write("<title>")
  Response.Write("<![CDATA[" & (objRS.Fields.Item("BlogHeadline").Value) & "]]>")
  Response.Write("</title>")
  Response.Write("<link>" & (rsSiteDetails.Fields.Item("blogURL").Value) & "/permalink.asp?id=" & (objRS.Fields.Item("BlogID").Value) & "</link>")
  Response.Write("<description>")
  Response.Write("<![CDATA[" & (objRS.Fields.Item("BlogHTML").Value) & "]]>")
  Response.Write("</description>")
  Response.Write("</item>")
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  objRS.MoveNext()
Wend
  Response.Write("</channel>")
  Response.Write("</rss>")
objRS.Close()
Set objRS = Nothing
rsSiteDetails.Close()
Set rsSiteDetails = Nothing
%>
