<!--#include file = "../Inc/Syscode.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY topMargin=0 leftmargin="5" marginheight="0">
<LINK href="../style.css" rel=stylesheet type=text/css>
<script language="javascript">
function del_class_1(id){
if (confirm("真的要删除这个一级分类吗?")) 
window.location.href="Admin_class_ok.asp?action=del_class_1&id="+id+""
} 
</script> <script language="javascript">
function confirmdel(id){
if (confirm("真的要删除这个二级分类吗?")) 
window.location.href="Admin_class_ok.asp?action=del_class_2&id="+id+""
} 
</script> <SCRIPT language=javascript>
function FORM1_onsubmit()
{
if(document.FORM1.class_name.value.length<1)
{
alert("您必须输入商品大类名称!");
document.FORM1.class_name.focus();
return false;
}
}
</SCRIPT> </HEAD> <%
set rs_class_1=server.createobject("adodb.recordset")
sqltext1="select * from Class_1 "
rs_class_1.open sqltext1,conn,1,1
%> <CENTER> <table cellpadding=0 cellspacing=0 border=0 width=100% align=center> 
<tr> <td> <table cellpadding=3 cellspacing=1 border=0 width=100%> <tr> <td width="100%" valign=top align=center> 
<TABLE border=0 borderColor=#ffffff cellPadding=0 cellSpacing=1> <TR><TD align="left">说明：删除商品大类前请先删除商品小类 
</TD></TR> <TR> <TD width="100%" colspan="2" height="77"> <%
While Not rs_class_1.EOF
set rs_class_2=server.createobject("adodb.recordset")
sqltext2="select * from Class_2 where sortid="& rs_class_1("sortid")&""
rs_class_2.open sqltext2,conn,1,1 
%> <TABLE border=0 cellPadding=0 cellSpacing=0 width=100%> <TR> <TD height="28" align="left">&nbsp; 
<IMG SRC='../images/workplan.gif' BORDER=0 width="30" height="16">【<%=rs_class_1("sort")%>】<a href=Admin_Class_Ok.asp?action=edit_class_1&id=<%=rs_class_1("sortid")%>><font color=#ff0000>修改</font></a>&nbsp;&nbsp;<a href='javascript:del_class_1(<%=rs_class_1("sortid")%>)'><font color=#ff0000>删除</font></a></TD></TR> 
<%j=1%> <%While Not rs_class_2.EOF%> <TR height=16> <td>&nbsp; <IMG SRC='../images/blank.gif' BORDER=0><font color="#B1DB99"><%=rs_class_2("typeName")%></font></a>〖<a href=Admin_Class_Ok.asp?action=edit_class_2&id=<%=rs_class_2("typeid")%>><font color=#ff0000>修改</font></a>&nbsp;&nbsp;<a href='javascript:confirmdel(<%=rs_class_2("typeid")%>)'><font color=#ff0000>删除</font></a>〗</TD><TD align=center width=100></TD></TR> 
<% 
rs_class_2.MoveNext
j=j+1
Wend
rs_class_2.close
%> </TABLE><%
rs_class_1.MoveNext
Wend
rs_class_1.close
conn.close
%> </TD></TR> </TABLE></TD></TR> </TABLE></TD></TR> </TABLE>