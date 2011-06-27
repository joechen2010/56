<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
Dim gsid
gsid=SafeRequest("gsid",1)
if gsid="" then
    call WriteErrMsg(errmsg)
end if

set rs_q=server.createobject("adodb.recordset")
sql="select * from qyml where id="&gsid&""
rs_q.open sql,conn,1,1
if rs_q.eof and rs_q.bof then
    call WriteErrMsg(errmsg)
else
    if rs_q("flag")=0 then
	    Response.Redirect "aboutus.asp?gsid="&gsid
	else
	   if rs_q("url")<>"" then
           Response.Redirect rs_q("url")
       else
           Response.Redirect "gold01/index.asp?gsid="&gsid
	   End IF
	end if
end if
 %>