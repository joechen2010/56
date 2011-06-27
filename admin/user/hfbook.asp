<%@ codepage ="936" %>
<%
if instr(session("flag"),"01")=6 then
response.redirect "../login.asp"
response.end
end if
%>
<!--#include file="../../Inc/Conn.asp" -->

<%
dim action
action=request("action")
if action="add" then
call add()
else
call main()
end if

sub main()
Dim GbId
GbId = request.QueryString("Id")
set rsu=server.createobject("adodb.recordset")
sql="select Id,Title from GuestBook where id="&GbId
rsu.open sql,conn,1,1 
%>
<link rel="stylesheet" type="text/css" href="../style.css">
<STYLE type=text/css>
<!--
td { font-family: "宋体"; font-size: 9pt; line-height: 14pt; }
-->
</STYLE> 
<p></p><div align="center"><center>
    <table width="80%" border="0" cellpadding="6" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111">
      <tr> <td height="20" bgcolor="#E8F4FF" width="80%"> <div align="center"><strong> 
<font color="#FF0000" style="font-size: 10.5pt">回 复 留 言</font></strong></div></td></tr> 
<tr> <Form language=javascript name="Gbhf" method="POST" action="hfbook.asp?action=add" onsubmit="return Gbhf_onsubmit()"> 
<td height="56" valign="top" width="80%"> <div align="center"> <center> <table width="100%" border="1" cellspacing="1" bgcolor="F4F4F4" bordercolorlight="#E8F4FF" bordercolordark="#E8F4FF" style="border-collapse: collapse" bordercolor="#111111"> 
<tr> <td height="30" align="right" width="128" bgcolor="#FBFDFF">主题：</td><td height="30" width="524" bgcolor="#FBFDFF"><font color="#FF0000"><%=rsu("Title")%></font> 
<input type="hidden" name="G_id" value="<%=rsu("Id")%>"></td></tr> <tr> <td width="128" height="30" align="right" bgcolor="#FBFDFF">姓名：</td><td width="524" height="30" bgcolor="#FBFDFF"> 
<% IF not(Session("KEY")="1" or session("KEY")="2") THEN %> <input name="G_Name" type="text" maxlength="20" size="20"> 
<font color="#FF0000">*</font> 请写上您的姓名。 <%else%> <input name="G_Name" type="text" maxlength="20" value="<%=Session("UserName")%>" size="20"> 
<%end if%> </td></tr> <tr> <td width="128" height="30" align="right" bgcolor="#FBFDFF">邮件：</td><td height="30" width="524" bgcolor="#FBFDFF"> 
<input name="G_Email" type="text" id="G_Email" size="29" maxlength="50"> （请填上您的Email地址。） 
</td></tr> <tr> <td width="128" align="right" bgcolor="#FBFDFF">内容：</td><td width="524" bgcolor="#FBFDFF"> 
<textarea name="G_Content" cols="55" rows="8" id="G_Content"></textarea> <font color="#FF0000">*</font></td></tr> 
<tr align="center"> <td height="30" colspan="2" width="653" bgcolor="#FBFDFF"> 
<input type="submit" name="Submit" value="回 复"> <input type="button" value="返 回" class="smallInput" onclick=javascript:history.go(-1)> 
</td></tr> </table></center></div></td></Form></tr> <tr> <td height="20" valign="top" bgcolor="#E8F4FF" width="80%">　</td></tr> 
</table></center></div><SCRIPT language=javascript id=QzeNet>
<!--
function Gbhf_onsubmit() 
{
if(document.Gbhf.G_Name.value.length<1)
{
alert("您必须输入您的姓名!");
document.Gbhf.G_Name.focus();
return false;
}
if(document.Gbhf.G_Content.value.length<1)
{
alert("您必须输入您的留言内容!");
document.Gbhf.G_Content.focus();
return false;
}
if(document.Gbhf.G_Content.value.length>300)
{
alert("您输入留言内容超过300字!");
document.Gbhf.G_Content.focus();
return false;
}
}
//-->
</SCRIPT> <% 
end sub

sub add()
Dim Gid,Gname,Gemail,Gcontent
GId = request.form("G_Id")
Gname = Trim(request.form("G_Name"))
Gemail = Trim(request.form("G_Email"))
Gcontent = request.form("G_Content")
'HTML支持

set rs=server.createobject("adodb.recordset")
sqltext="select * from hfbook"
rs.open sqltext,conn,3,3
'添加一条留言到数据库
rs.addnew
rs("GId")=GId
rs("GName")=Gname
if Gemail="" then
rs("GEmail")="@"
else
rs("GEmail")=Gemail
end if
rs("GContent")=htmlencode(Gcontent)
if Session("GbAdmin")="" then
rs("IsA")=false
else
rs("IsA")=True
end if
rs("GTime")=cstr(now())
rs.update
rs.close
conn.close
Response.Redirect "khbook.asp"
end sub

function htmlencode(str)
htmlencode=server.htmlencode(str)
htmlencode=replace(replace(htmlencode,chr(13),"<br>"),"'","’")
end function
%>