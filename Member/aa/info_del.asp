<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<%
id=request("id")
set rs=server.createobject("adodb.recordset")
sql="select * from info where info_id="&id
rs.open sql,conn,3,2
if not rs.eof then
rs.delete
rs.update
rs.close
end if
response.write "<script>alert('数据删除成功！');top.location.href='info_manage.asp';</script>"
response.end
%>