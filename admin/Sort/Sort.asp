<!--#include file = "../Inc/Syscode.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../style.css" rel=stylesheet type=text/css> 
<script language="javascript">
function del_class_1(id){
if (confirm("真的要删除这个大类吗?")) 
window.location.href="Admin_class_ok.asp?action=del_class_1&id="+id+""
} 
</script> </HEAD> 
<BODY topMargin=0 leftmargin="5" marginheight="0" bgcolor="#e1eefd">
<%
set rssort=server.createobject("adodb.recordset")
sql="select * from Class_1 "
rssort.open sql,conn,1,1
%> <CENTER> <table width="620" border="0" cellspacing="0" cellpadding="3"> <TR> 
<TD height=40>行业大类管理（点击相应的分类进行操作）-- <a href="Admin_Class_Ok.asp?action=add_class_1">添加大类</a></td></tr> 
<tr> <td valign=top> <TABLE border="0" width="100%" cellpadding="0"> <TR> <%j=1
do while not rssort.eof
%> <td width="28%" height="25" align="left"><p style="line-height: 150%"><IMG src="../images/dir.gif"> 
<a href="Type.asp?sortid=<%=rssort("sortid")%>"><%=rssort("sort")%></a> 〖<a href=Admin_Class_Ok.asp?action=edit_class_1&id=<%=rssort("sortid")%>><font color=#ff0000>修改</font></a>│<a href='javascript:del_class_1(<%=rssort("sortid")%>)'><font color=#ff0000>删除</font></a>〗</td><%if j mod 3 = 0 then %></tr><tr ><%end if%> 
<% rssort.movenext
j=j+1
loop
rssort.close
set rssort=nothing
%> </TABLE></TD></TR> </TBODY> </TABLE></body>