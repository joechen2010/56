<!--#include file = "../Inc/Syscode.asp"-->
<%
if not isEmpty(request("sortid")) then
sortid=request("sortid")
else
Response.Redirect "type.asp"
Response.End
end if
%>
<script language="javascript">
function confirmdel(id){
if (confirm("���Ҫɾ�����������?")) 
window.location.href="Admin_class_ok.asp?action=del_class_2&id="+id+""
} 
</script>
<%
dim rssort,sortsql
dim body
dim rstype,typesql
dim sortname
dim sortid
dim totle
Set rssort= Server.CreateObject("ADODB.Recordset") 
sortsql="select * from class_1 where sortid="&sortid&"" 
rssort.open sortsql,conn,1,1
sortname=rssort("sort")
If rssort.eof and rssort.bof then 
response.write"<SCRIPT language=JavaScript>alert('�Բ��𣬲�������!');"
response.write"this.location.href='javascript:history.back();'</SCRIPT>"
End if
rssort.close
set rssort=nothing
set rstype=server.createobject("adodb.recordset") 
typesql="select * from class_2 where sortid="&sortid&" order by typeid asc" 
rstype.open typesql,conn,1,1 
%>
<html>
<LINK href="../style.css" rel=stylesheet type=text/css> 

<BODY topMargin=0 leftmargin="5" marginheight="0" bgcolor="#e1eefd">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%"> 
<TBODY> <TR !bgcolor="#f5f5f5"> <TD class=f2>��ҵС�����-- <a href="Admin_Class_Ok.asp?action=add_class_2">�����ҵС��</a></TD></TR> 
</TBODY> </TABLE><TABLE align=center border=0 cellPadding=3 cellSpacing=0 width=100%> 
<TBODY> <TR> <TD align=middle vAlign=top> <TABLE border="0" width="98%" cellpadding="0"> 
<TR height=22> <%j=1
do while not rstype.eof%> <td width="33%" align="left" height=24> <IMG src="../images/list.gif"> 
<%=rstype("typename")%>��<a href=Admin_Class_Ok.asp?action=edit_class_2&id=<%=rstype("typeid")%>><font color=#ff0000>�޸�</font></a>&nbsp;&nbsp;<a href='javascript:confirmdel(<%=rstype("typeid")%>)'><font color=#ff0000>ɾ��</font></a>�� 
</td><%if j mod 3 = 0 then %></tr><tr ><%end if%> <% rstype.movenext 
j=j+1 
loop 
rstype.close
set rstype=nothing 
%> </tr> </TABLE></TD></TR> </TABLE></BODY></HTML>