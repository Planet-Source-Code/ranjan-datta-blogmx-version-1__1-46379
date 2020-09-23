// JavaScript Document
<!-- Begin
function AddText(form, Action){
var AddTxt="";
var txt="";
if(Action==1){  
txt=prompt("Text for the level 2 header.","Text");      
if(txt!=null)           
AddTxt="<h2>"+txt+"</h2>\r\n";  
}
if(Action==2){  
txt=prompt("Text for the level 3 header.","Text");      
if(txt!=null)           
AddTxt="<h3>"+txt+"</h3>\r\n";  
}
if(Action==3) {  
txt=prompt("Text to be made BOLD.","Text");     
if(txt!=null)           
AddTxt="<strong>"+txt+"</strong>";        
}
if(Action==4) {  
txt=prompt("Text to be italicized","Text");     
if(txt!=null)           
AddTxt="<em>"+txt+"</em>";        
}
if(Action==5) AddTxt="\r\n<p></p>";
if(Action==6) AddTxt="<br />\r\n";
if(Action==7) AddTxt="<hr />\r\n";
if(Action==8) {  
txt=prompt("URL for the link.","http://");      
if(txt!=null){          
AddTxt="<a href=\""+txt+"\">";              
txt=prompt("Text to be show for the link","Text");              
AddTxt+=txt+"</a>\r\n";         
   }
}
if(Action==9) { 
txt=prompt("URL for image","URL");    
talt=prompt("Alt for image", "empty");
AddTxt="<img src=\""+txt+"\" alt=\""+talt+"\" />\r\n"; 
}
if(Action==10) {  
txt=prompt("Abbr","Abbr");      
txtTitle=prompt("Title","Title");              
AddTxt+="<abbr title=\""+txtTitle+"\">"+txt+"</abbr>";         
}
if(Action==11) {  
AddTxt="\r\n<ul>\r\n<li>Add Text Here</li>\r\n<li>Add Text Here</li>\r\n<li>Add Text Here</li>\r\n<li>Add Text Here</li>\r\n</ul>\r\n";         
}
form.Text1.value+=AddTxt;
}
// End -->