<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from message where id="&request("id")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
    	Response.Write "<script language=javascript>alert('��¼�����ڣ�\n\nϵͳ���Զ�����ǰһҳ��...');history.back();</script>"
else
    conn.execute("delete from message where id="&request("id"))
	Response.Write "<script language=javascript>alert('���ݳɹ���\n\nϵͳ���Զ�����ǰһҳ��...');history.back();</script>"
	'Response.End
end if
%>
