<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../member/check.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from message where id="&request("id")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
    	Response.Write "<script language=javascript>alert('��¼�����ڣ�\n\nϵͳ���Զ�����ǰһҳ��...');history.back();</script>"
else
    conn.execute("delete from message where id="&request("id"))
response.write "<script>alert('��Ϣɾ���ɹ���');top.location.href='site.asp?infoid="&request("infoid")&"&user="&request("user")&"&action="&request("action")&"';</script>"	'Response.End
end if
%>
