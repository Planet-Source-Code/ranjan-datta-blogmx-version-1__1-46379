<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "../default.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
 
<ul>
<li><a href="../default.asp">Blog Page</a></li>
<li><a href="main.asp">Main Admin Page</a></li>
<li><a href="add_category.asp">Add Category</a></li>
<li><a href="add_blog.asp">Add Blog</a> </li>
<li><a href="add_links.asp">Add / Delete Link</a> </li>
<li><a href="change_user.asp">Change User</a> </li>
<li><a href="approve_comments.asp">Approve Comments</a> </li>
<li><a href="blog_config.asp">Blog Configuration</a></li>
<li><a href="<%= MM_Logout %>">Logout</a></li>
</ul>
