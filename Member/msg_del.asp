<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from message where id="&request("id")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
    	Response.Write "<script language=javascript>alert('记录不存在！\n\n系统将自动返回前一页面...');history.back();</script>"
else
    conn.execute("delete from message where id="&request("id"))
	Response.Write "<script language=javascript>alert('数据成功！\n\n系统将自动返回前一页面...');history.back();</script>"
	'Response.End
end if
%>
