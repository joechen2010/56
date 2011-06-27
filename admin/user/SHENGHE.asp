<%@ codepage ="936" %>
<!--#include file="../../Inc/Conn.asp" -->
<%
id=request("id")
action=request("action")
if action="juflag" then
    conn.execute("update qyml set Uflag=1 where id="&id)
	response.redirect "default.asp?id="&request("id")&"&page="&request("Page")&""
elseif action="cuflag" then
    conn.execute("update qyml set Uflag=0 where id="&id)
	response.redirect "check.asp?id="&request("id")&"&page="&request("Page")&""
end if
%>