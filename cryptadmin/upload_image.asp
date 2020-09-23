<HTML>
<HEAD>
<!--#include file="clsUpload.asp"-->
<script type="text/javascript" language="JavaScript">
function doCopy() {

Copied = document.form1.mypath.createTextRange();
Copied.execCommand("Copy");
}

</script>
<link href="../styles/adminV5.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>
<FORM name="uploadFrm" ACTION = "upload_image.asp" ENCTYPE="multipart/form-data" METHOD="POST">

<p>Image Name: 
<label for="txtName"></label>
<INPUT NAME = "Demo" class="txtBox">
</INPUT>
</p>
<p>File Name: 
<INPUT NAME="txtFile" TYPE=FILE class="txtBox">
<INPUT NAME="cmdSubmit" TYPE = "SUBMIT" class="txtBox" VALUE="Upload">
</p>
</FORM>

<%




set o = new clsUpload
if o.Exists("cmdSubmit") then

'get client file name without path
sFileSplit = split(o.FileNameOf("txtFile"), "\")
sFile = sFileSplit(Ubound(sFileSplit))

o.FileInputName = "txtFile"
o.FileFullPath = Server.MapPath("uploaded_images") & "\" & sFile
o.save
 
	if o.Error = "" then
	%>
	
<p>Success. File <%= o.ValueOf("Demo") %> saved to  <%=o.FileFullPath%></p>
<form name="form1" method="post" action="">
<p>
<input name="mypath" type="text" class="txtBox" id="mypath" value="<img src='<%= o.FileFullPath %>' alt='<%= o.ValueOf("Demo") %>'>" size="35">
<a href="javascript:;" onClick="doCopy()">copy</a> </p>
<p>Click Copy to copy path and paste in blog text area document</p>
<p><a href="javascript:window.close()">Close Window</a></p>
</form>
<%
	else
	%>
<p>	Failed due to the following error:  <%=o.Error%></p>
 
<%
	end if

end if
set o = nothing
%>
</BODY>
</HTML>
