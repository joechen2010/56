<!--#include file = "../Inc/Syscode.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY topMargin=0 leftmargin="0" marginheight="0">
<LINK href="../style.css" rel=stylesheet type=text/css>  <%
dim body
call main()
conn.close
set conn=nothing
sub main()
%>  <table cellpadding=0 cellspacing=0 border=0 width=100% align=center> 
<tr> <td> <table cellpadding=4 cellspacing=1 border=0 width=100%> <tr> <td width="100%" valign=top><font color="<%=TableContentColor%>"> 
<%
if request("action") ="add_class_1" then 
call add_class_1()
elseif request("action")="add_class_1_ok" then
call add_class_1_ok()
elseif request("action")="add_class_2" then
call add_class_2()
elseif request("action")="add_class_2_name" then
call add_class_2_name()
elseif request("action")="add_class_2_ok" then
call add_class_2_ok()
elseif request("action")="edit_class_1" then
call edit_class_1()
elseif request("action")="edit_class_2" then
call edit_class_2()
elseif request("action")="savedit_1" then
call savedit_1()
elseif request("action")="savedit_2" then
call savedit_2()
elseif request("action")="del_class_1" then
call del_class_1()
elseif request("action")="del_class_2" then
call del_class_2()
else
call linkinfo()
end if
%> <p><%=body%></p></font> </td></tr> </table></td></tr> </table><%
end sub
sub add_class_1()
%> <SCRIPT language=javascript>
function FORM1_onsubmit()
{
if(document.FORM1.class_name.value.length<1)
{
alert("�����������������!");
document.FORM1.class_name.focus();
return false;
}
}
</SCRIPT> <FORM name=FORM1 onSubmit="return FORM1_onsubmit()" action=Admin_Class_Ok.asp?action=add_class_1_ok method=post> 
<TABLE cellSpacing=1 cellPadding=4 width=600> <TBODY> <TR vAlign=middle> <TD> 
<TABLE cellPadding=4 cellSpacing=0 width=100% border="0"> <TBODY> <TR><TD colSpan=2 height="28" bgcolor=#e8f4ff>��ҵ�������--</TD></TR> 
<TR> <TD width=40% height=25><b>��ҵ��������</b></TD><TD width=60% height=25><INPUT maxLength=16 size=25 name=class_name></TD></TR> 
<TR> <TD colSpan=2 height="27" align=center bgcolor=#e8f4ff> <INPUT type=submit value='ȷ ��' name=Submit2> 
</TD></TR> </TBODY> </TABLE></TD></TR> </TABLE></FORM><%
end sub
sub add_class_2()
set rs=server.createobject("adodb.recordset")
sqltext2="select * from Class_1 "
rs.open sqltext2,conn,1,1
%> <FORM name=FORM1 action=Admin_Class_Ok.asp?action=add_class_2_name method=post> 
<br><br> <TABLE cellPadding=4 cellSpacing=0 width=600 border="0"> <TBODY> <TR> 
<TD colSpan=2 height="32" bgcolor=#e8f4ff>��ҵС�����--(��һ��)</TD></TR> <TR> <TD width=40% height=25><b>������ҵ����</b><br>��ѡ�����Ҫ��ӵ���ҵС�������Ĵ��ࡣ</TD><TD width=60% height=25><select size="1" name="class_1_id" style="font-size: 9pt"> 
<%While Not rs.EOF%> <option value=<%=rs("sortid")%>><% =rs("sort")%></option> 
<%
rs.MoveNext
Wend
rs.close
%> </select></TD></TR> <TR> <TD colSpan=2 height="27" align=center bgcolor=#e8f4ff> 
<INPUT class=smallInput type=submit size=3 value='��һ��' name=Submit2> </TD></TR> 
</TBODY> </TABLE></form><%
end sub
sub add_class_2_name()
set rs=server.createobject("adodb.recordset")
sqltext2="select * from Class_1 where sortid="+Cstr(request("class_1_id"))
rs.open sqltext2,conn,1,1
%> <SCRIPT language=javascript>
function FORM1_onsubmit()
{
if(document.FORM1.class_2_name.value.length<1)
{
alert("������������ҵС������!");
document.FORM1.class_2_name.focus();
return false;
}
}
</SCRIPT> <FORM name=FORM1 onSubmit="return FORM1_onsubmit()" action=Admin_Class_Ok.asp?action=add_class_2_ok method=post> 
<INPUT type=hidden value="<%=request.form("class_1_id")%>" name="class_1_id"> 
<TABLE cellPadding=4 cellSpacing=0 width=100% border="0"> <TBODY> <TR> <TD colSpan=2 height="32" bgcolor=#e8f4ff>��ҵС�����--(�ڶ���)</TD></TR> 
<TR> <TD width=40% height=25><b>������ҵ����</b><br>������ʾ�������ѡ�����ҵ��������ƣ������޸�������һ����</TD><TD width=60% height=25><%=rs("sort")%></TD></TR> 
<TR> <TD width=40% height=25><b>��ҵС������</b></TD><TD width=60% height=25><input type="text" name="class_2_name" size="18"></TD></TR> 
<TR> <TD colSpan=2 height="27" align="center" bgcolor=#e8f4ff><input type="button" value="��һ��" name="B4" onClick="javascript:window.history.go(-1)"> 
<INPUT type=submit size=3 value='ȷ&nbsp;&nbsp;��' name=Submit2></TD></TR> </TBODY> 
</TABLE>�� </form><%
end sub
sub add_class_2_ok()
set rs=server.createobject("adodb.recordset")
sqltext="select * from Class_2 where typename='"&request.form("class_2_name")&"' and sortid=" & request.form("class_1_id")&""
rs.open sqltext,conn,1,1
'�������ݿ⣬�����������Ƿ��Ѿ�����
if rs.recordcount >= 1 then 
if rs("typename")=Lcase(request.form("class_2_name")) then
response.write"<SCRIPT language=JavaScript>alert('�˶��������Ѿ����ڣ���ѡ����������!');"
response.write"this.location.href='javascript:history.back();'</SCRIPT>"
rs.close
end if
else
set rs=server.createobject("adodb.recordset")
sqltext="select * from Class_2"
rs.open sqltext,conn,3,3
'���һ���������ൽ���ݿ�
rs.addnew
rs("typename")=request.form("class_2_name")
rs("sortid")=request.form("class_1_id")
rs.update
rs.close
set rs=nothing
response.redirect "sort.asp"
end if
end sub
sub edit_class_1()
set rs=server.createobject("adodb.recordset")
sqltext="select * from Class_1 where sortid="+Cstr(request("id"))
rs.open sqltext,conn,1,1
%> <form name=form1 action="Admin_Class_Ok.asp?action=savedit_1" method=post> 
<input type=hidden name=id value="<%=Cstr(request("id"))%>"> <br><br> 
<table width="100%" border="0" cellspacing="4" cellpadding="0"> <TR> <TD colSpan=2 height="28" bgcolor="#e8f4ff">��ҵ�����޸�</TD></TR> 
<tr> <td width="40%"><b>��������</b></td><td width="60%"> <input type="text" name="class_1_name" size='40' value=<%=rs("sort")%>> 
</td></tr> <tr> <td height="30" colspan="2" bgcolor="#e8f4ff" align="center"> 
<input type="submit" name="Submit" value="�޸�"> </td></tr> </table></form><%
rs.close
set rs=nothing
end sub
sub edit_class_2()
set rs=server.createobject("adodb.recordset")
sqltext="select * from Class_2 where typeid="+Cstr(request("id"))
rs.open sqltext,conn,1,1
set rs2=server.createobject("adodb.recordset")
sqltext2="select * from Class_1 where sortid="+Cstr(rs("sortid"))
rs2.open sqltext2,conn,1,1
%> <form name=form1 action=Admin_Class_Ok.asp?action=savedit_2 method=post> <input type=hidden name=id value=<%=Cstr(request("id"))%>>
<br><br> <table width="100%" border="0" cellspacing="4" cellpadding="0"> <TR> 
<TD colSpan=2 height="28" bgcolor="#e8f4ff">��Ʒ���������޸�</TD></TR> <tr> <td width="40%"><b>һ����������</b></td><td width="60%"><%=rs2("sort")%></td></tr> 
<tr> <td width="40%"><b>������������</b></td><td width="60%"> <input type="text" name="class_2_name" size='40' value=<%=rs("typename")%>> 
</td></tr> <tr> <td height="30" colspan="2" bgcolor="#e8f4ff" align="center"> 
<input type="submit" name="Submit" value="�޸�"> </td></tr> </table></form><%
rs.close
rs2.close
set rs=nothing
set rs2=nothing
end sub
sub savedit_1()
set rs= server.createobject ("adodb.recordset")
sql ="select * from class_1 where sortid="+Cstr(request("id"))
rs.Open sql,conn,1,3
if rs.eof and rs.bof then
body=body+"<br>"+"�������ݿ��������û���ҵ�������Ϣ��"
else
rs("sort")=Trim(Request.Form ("class_1_name"))
rs.Update
rs.close
set rs=nothing
end if 
response.redirect "sort.asp"
end sub
sub savedit_2()
set rs= server.createobject ("adodb.recordset")
sql ="select * from class_2 where typeid="+Cstr(request("id"))
rs.Open sql,conn,1,3
if rs.eof and rs.bof then
body=body+"<br>"+"�������ݿ��������û���ҵ�������Ϣ"
else
rs("typename")=Trim(Request.Form ("class_2_name"))
rs.Update
rs.close
set rs=nothing
end if 
response.redirect "sort.asp"
end sub
sub del_class_1()
set rs1=server.createobject("adodb.recordset")
sqltext1="select * from class_2 where sortid="+Cstr(request("id"))
rs1.open sqltext1,conn,1,1
If rs1.eof and rs1.bof then 
set rs=server.createobject("adodb.recordset")
sqltext="delete from class_1 where sortid="+Cstr(request("id"))
rs.open sqltext,conn,3,3
rs.close
set rs=nothing
response.redirect "sort.asp"
else
rs1.close
set rs1=nothing
response.write"<SCRIPT language=JavaScript>alert('�Բ��𣬱�������滹��С�࣡');"
response.write"this.location.href='javascript:history.back();'</SCRIPT>"
End if
end sub
sub del_class_2()
set rs=server.createobject("adodb.recordset")
sqltext="delete from class_2 where typeId="+Cstr(request("id"))
rs.open sqltext,conn,3,3
rs.close
set rs=nothing
response.redirect "sort.asp"
end sub
sub add_class_1_ok()
set rs=server.createobject("adodb.recordset")
sqltext="select * from Class_1 where sort='"&request.form("class_name")&"'"
rs.open sqltext,conn,1,1
'�������ݿ⣬�����Ʒ�����Ƿ��Ѿ�����
if rs.recordcount >= 1 then 
if rs("Class_1_name")=request.form("class_name") then
response.write"<SCRIPT language=JavaScript>alert('������Ѿ����ڣ���ѡ���������ƣ�');"
response.write"this.location.href='admin_class_ok.asp?action=add_class_1';</SCRIPT>"
end if
else
set rs=server.createobject("adodb.recordset")
sqltext="select * from Class_1"
rs.open sqltext,conn,3,3
rs.addnew
rs("sort")=request.form("class_name")
rs.update
rs.close
set rs=nothing
response.redirect "sort.asp"
end if
end sub
%> 
</body>
</html>