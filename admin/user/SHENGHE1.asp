<%@ codepage ="936" %>
<!--#include file="../../Inc/Conn.asp" -->
<%
id=request("id")
action=request("action")
if action="jflag" then
    conn.execute("update qyml set flag=1,begindate='"&date()&"',enddate='"&date()+365&"' where id="&id)
	response.redirect "charge.asp?id="&request("id")&"&page="&request("Page")&""
elseif action="cflag" then
    conn.execute("update qyml set flag=0 where id="&id)
	response.redirect "payment.asp?id="&request("id")&"&page="&request("Page")&""
end if

%>