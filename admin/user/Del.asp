<%@ codepage ="936" %>
<!--#include file="../../Inc/Conn.asp" -->
<%
dim sql,rs,Id,gid
Id=request.querystring("id")

conn.execute("delete from guestbook where id="&Id)
conn.execute("delete from hfbook where gid="&id)
response.Redirect"khbook.asp"

%>