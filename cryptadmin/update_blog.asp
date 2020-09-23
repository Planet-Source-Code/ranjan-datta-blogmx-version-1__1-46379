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
  MM_editTable = "tblBlog"
  MM_editColumn = "BlogID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "main.asp"
  MM_fieldsStr  = "txtBlogHeading|value|chkPublish|value|txtCat|value|Text1|value"
  MM_columnsStr = "BlogHeadline|',none,''|BlogIncluded|none,1,0|CatID|none,none,NULL|BlogHTML|',none,''"

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
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="default.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
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
<%
Dim rsBlog__MMColParam
rsBlog__MMColParam = "1"
If (Request.QueryString("PassID") <> "") Then 
  rsBlog__MMColParam = Request.QueryString("PassID")
End If
%>
<%
Dim rsBlog
Dim rsBlog_numRows

Set rsBlog = Server.CreateObject("ADODB.Recordset")
rsBlog.ActiveConnection = MM_dreamConn_STRING
rsBlog.Source = "SELECT * FROM tblBlog WHERE BlogID = " + Replace(rsBlog__MMColParam, "'", "''") + ""
rsBlog.CursorType = 0
rsBlog.CursorLocation = 2
rsBlog.LockType = 1
rsBlog.Open()

rsBlog_numRows = 0
%>
<html>
<head>
<title>blogMX :: Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script  type="text/JavaScript" src="editorjs/editor.js"></script>
<style type="text/css">
<!--
-->
</style>
<style type="text/css">
<!--
-->
</style>
<script  type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function YY_checkform() { //v4.65
//copyright (c)1998,2002 Yaromat.com
  var args = YY_checkform.arguments; var myDot=true; var myV=''; var myErr='';var addErr=false;var myReq;
  for (var i=1; i<args.length;i=i+4){
    if (args[i+1].charAt(0)=='#'){myReq=true; args[i+1]=args[i+1].substring(1);}else{myReq=false}
    var myObj = MM_findObj(args[i].replace(/\[\d+\]/ig,""));
    myV=myObj.value;
    if (myObj.type=='text'||myObj.type=='password'||myObj.type=='hidden'){
      if (myReq&&myObj.value.length==0){addErr=true}
      if ((myV.length>0)&&(args[i+2]==1)){ //fromto
        var myMa=args[i+1].split('_');if(isNaN(parseInt(myV))||myV<myMa[0]/1||myV > myMa[1]/1){addErr=true}
      } else if ((myV.length>0)&&(args[i+2]==2)){
          var rx=new RegExp("^[\\w\.=-]+@[\\w\\.-]+\\.[a-z]{2,4}$");if(!rx.test(myV))addErr=true;
      } else if ((myV.length>0)&&(args[i+2]==3)){ // date
        var myMa=args[i+1].split("#"); var myAt=myV.match(myMa[0]);
        if(myAt){
          var myD=(myAt[myMa[1]])?myAt[myMa[1]]:1; var myM=myAt[myMa[2]]-1; var myY=myAt[myMa[3]];
          var myDate=new Date(myY,myM,myD);
          if(myDate.getFullYear()!=myY||myDate.getDate()!=myD||myDate.getMonth()!=myM){addErr=true};
        }else{addErr=true}
      } else if ((myV.length>0)&&(args[i+2]==4)){ // time
        var myMa=args[i+1].split("#"); var myAt=myV.match(myMa[0]);if(!myAt){addErr=true}
      } else if (myV.length>0&&args[i+2]==5){ // check this 2
            var myObj1 = MM_findObj(args[i+1].replace(/\[\d+\]/ig,""));
            if(myObj1.length)myObj1=myObj1[args[i+1].replace(/(.*\[)|(\].*)/ig,"")];
            if(!myObj1.checked){addErr=true}
      } else if (myV.length>0&&args[i+2]==6){ // the same
            var myObj1 = MM_findObj(args[i+1]);
            if(myV!=myObj1.value){addErr=true}
      }
    } else
    if (!myObj.type&&myObj.length>0&&myObj[0].type=='radio'){
          var myTest = args[i].match(/(.*)\[(\d+)\].*/i);
          var myObj1=(myObj.length>1)?myObj[myTest[2]]:myObj;
      if (args[i+2]==1&&myObj1&&myObj1.checked&&MM_findObj(args[i+1]).value.length/1==0){addErr=true}
      if (args[i+2]==2){
        var myDot=false;
        for(var j=0;j<myObj.length;j++){myDot=myDot||myObj[j].checked}
        if(!myDot){myErr+='* ' +args[i+3]+'\n'}
      }
    } else if (myObj.type=='checkbox'){
      if(args[i+2]==1&&myObj.checked==false){addErr=true}
      if(args[i+2]==2&&myObj.checked&&MM_findObj(args[i+1]).value.length/1==0){addErr=true}
    } else if (myObj.type=='select-one'||myObj.type=='select-multiple'){
      if(args[i+2]==1&&myObj.selectedIndex/1==0){addErr=true}
    }else if (myObj.type=='textarea'){
      if(myV.length<args[i+1]){addErr=true}
    }
    if (addErr){myErr+='* '+args[i+3]+'\n'; addErr=false}
  }
  if (myErr!=''){alert('The required information is incomplete or contains errors:\t\t\t\t\t\n\n'+myErr)}
  document.MM_returnValue = (myErr=='');
}
//-->
</script>
</head>

<body>
<table width="100%"  border="0" cellspacing="2" cellpadding="3">
<tr>
<td>&nbsp;</td>
<td>&nbsp;</td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td><h1>Update Blog</h1></td>
</tr>
<tr>
<td>
<!--#include file="menu.asp" -->
</td>
<td><form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1" onSubmit="YY_checkform('form1','txtBlogHeading','#q','0','Enter Blog Heading','txtCat','#q','1','Select a Category','Text1','10','1','Blog something');return document.MM_returnValue">

<table width="100%"  border="0">

<tr> 
<td>Heading</td>
<td> <input name="txtBlogHeading" type="text" id="txtBlogHeading" value="<%=(rsBlog.Fields.Item("BlogHeadline").Value)%>" size="30"></td>
<td><input <%If rsBlog.Fields.Item("BlogIncluded").Value = -1 Then Response.Write("checked") : Response.Write("")%> name="chkPublish" type="checkbox" id="chkPublish"></td>
<td>Publish?</td>
<td>Category</td>
<td> <select name="txtCat" id="txtCat">
<option selected value="" <%If (Not isNull((rsBlog.Fields.Item("CatID").Value))) Then If ("" = CStr((rsBlog.Fields.Item("CatID").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>--Select 
One--</option>
<%
While (NOT rsCat.EOF)
%>
<option value="<%=(rsCat.Fields.Item("catID").Value)%>" <%If (Not isNull((rsBlog.Fields.Item("CatID").Value))) Then If (CStr(rsCat.Fields.Item("catID").Value) = CStr((rsBlog.Fields.Item("CatID").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsCat.Fields.Item("catName").Value)%></option>
<%
  rsCat.MoveNext()
Wend
If (rsCat.CursorType > 0) Then
  rsCat.MoveFirst
Else
  rsCat.Requery
End If
%>
</select></td>
<td><input type="submit" name="Submit" value="Update >>"></td>
</tr>
<tr> 
<td colspan="7"><table border="0" cellpadding="0" cellspacing="0">
<tr> 
<td id="tags"><strong>Allowed HTML Tags</strong>: <br> <a href="http://www.w3schools.com/tags/tag_a.asp">&lt;a&gt;</a><br> 
<a href="http://www.w3schools.com/tags/tag_acronym.asp">&lt;acronym&gt;</a><br> 
<a href="http://www.w3schools.com/tags/tag_abbr.asp">&lt;abbr&gt;</a><br> <a href="http://www.w3schools.com/tags/tag_blockquote.asp">&lt;blockquote&gt;</a><br> 
<a href="http://www.w3schools.com/tags/tag_br.asp">&lt;br&gt;</a><br> <a href="http://www.w3schools.com/tags/tag_phrase_elements.asp">&lt;em&gt;</a><br> 
<a href="http://www.w3schools.com/tags/tag_hn.asp">&lt;h2&gt;</a><br> <a href="http://www.w3schools.com/tags/tag_hn.asp">&lt;h3&gt;</a><br> 
<a href="http://www.w3schools.com/tags/tag_hr.asp">&lt;hr&gt;</a><br> <a href="http://www.w3schools.com/tags/tag_img.asp">&lt;img&gt;</a><br> 
<a href="http://www.w3schools.com/tags/tag_li.asp">&lt;li&gt;</a><br> <a href="http://www.w3schools.com/tags/tag_ol.asp">&lt;ol&gt;</a><br> 
<a href="http://www.w3schools.com/tags/tag_p.asp">&lt;p&gt;</a><br> <a href="http://www.w3schools.com/tags/tag_pre.asp">&lt;pre&gt;</a><br> 
<a href="http://www.w3schools.com/tags/tag_phrase_elements.asp">&lt;strong&gt;</a><br>
<a href="http://www.w3schools.com/tags/tag_ul.asp">&lt;ul&gt;</a> <a href="http://www.w3schools.com/tags/tag_phrase_elements.asp"><br>
&lt;code&gt;</a></td>
<td><textarea name="Text1" cols="75" rows="15" id="Text1"><%=(rsBlog.Fields.Item("BlogHTML").Value)%></textarea></td>
</tr>
<tr> 
<td>&nbsp;</td>
<td><input type="button" class="inputtags" onClick="AddText(this.form,1);" value="H2" /> 
<input type="button" value="H3" class="inputtags" onClick="AddText(this.form,2);" /> 
<input type="button" value="B" class="inputtags" onClick="AddText(this.form,3);" /> 
<input type="button" value="I" class="inputtags" onClick="AddText(this.form,4);" /> 
<input type="button" value="P" class="inputtags" onClick="AddText(this.form,5);" /> 
<input type="button" value="BR" class="inputtags" onClick="AddText(this.form,6);" /> 
<input type="button" value="HR" class="inputtags" onClick="AddText(this.form,7);" /> 
<input type="button" value="A" class="inputtags" onClick="AddText(this.form,8);" /> 
<input type="button" value="IMG" class="inputtags" onClick="AddText(this.form,9);" /> 
<input type="button" value="ABBR" class="inputtags" onClick="AddText(this.form,10);" /> 
<input type="button" value="UL" class="inputtags" onClick="AddText(this.form,11);" /></td>
</tr>
</table></td>
</tr>
</table>

<input type="hidden" name="MM_update" value="form1">
<input type="hidden" name="MM_recordId" value="<%= rsBlog.Fields.Item("BlogID").Value %>">
</form> </td>
</tr>
</table>
</body>
</html>
<%
rsCat.Close()
Set rsCat = Nothing
%>
<%
rsBlog.Close()
Set rsBlog = Nothing
%>
