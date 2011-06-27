<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../member/check.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from message where id="&request("id")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
    	Response.Write "<script language=javascript>alert('记录不存在！\n\n系统将自动返回前一页面...');history.back();</script>"
else
    conn.execute("delete from message where id="&request("id"))
response.write "<script>alert('信息删除成功！');top.location.href='site.asp?infoid="&request("infoid")&"&user="&request("user")&"&action="&request("action")&"';</script>"	'Response.End
end if
%>
