<%@ Language=VBScript codepage ="936" %>
<%
Response.Expires=true
Response.Buffer = true 
%>
<!--#include file="../inc/conn.asp"-->
<%
dim UsernameGet,PasswordGet
dim rs,sql
dim strComeUrl
if trim(request("strComeUrl"))="" then
	strComeUrl=Request.ServerVariables("HTTP_REFERER")
else
	strComeUrl=trim(request("strComeUrl"))
end if
UsernameGet=replace(trim(request("UserID")),"'","")
PasswordGet=replace(trim(Request("Pwd")),"'","")
sql="select * from qyml where User='"&UsernameGet&"' and Pass='"&PasswordGet&"' "
set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1
if rs.eof then
      rs.close
      response.write"<SCRIPT language=JavaScript>alert('您的用户名或密码有误!');"
      response.write"javascript:history.back(1)</SCRIPT>"
response.end
else 
conn.execute("update qyml set LastLoginIP='"&Request.ServerVariables("REMOTE_ADDR")&"',LastLoginTime='"&now&"',LoginTimes=LoginTimes+1 Where User='"&UsernameGet&"'") 
session("user")=UsernameGet
session("hydj")=rs("flag")
session("uflag")=rs("uflag")
session("Qymc")=rs("Qymc")
session("name")=rs("name")
session("email")=rs("email")
session("address")=rs("address")
session("post")=rs("post")
session("Phone")=rs("Phone1")&"-"&rs("Phone2")&"-"&rs("Phone3")
session("Fax")=rs("Fax1")&"-"&rs("Fax2")&"-"&rs("Fax3")
rs.close 
set rs=nothing 
conn.close
set conn=nothing 
end if
Response.Redirect strComeUrl
Response.End

%>
