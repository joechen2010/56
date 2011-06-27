<%@ Language=VBScript codepage ="936" %>
<%
Response.Expires=true
Response.Buffer = true 
%>
<!--#include file="../inc/conn.asp"-->
<%
dim UsernameGet,PasswordGet
dim rs,sql
UsernameGet=replace(trim(request("UserID")),"'","")
PasswordGet=replace(trim(Request("Pwd")),"'","")
'从哪来回哪去
if request("action")="login" then
url=request("url")
elseif request("action")="fz" then
url=request.ServerVariables("HTTP_REFERER")
end if
sql="select * from qyml where User='"&UsernameGet&"' and Pass='"&PasswordGet&"' and uflag=1"
set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,3
if rs.eof and rs.bof then
      rs.close
      response.write"<SCRIPT language=JavaScript>alert('您的用户名或密码有误,或没有通过审核!');"
      response.write"javascript:history.back(1)</SCRIPT>"
response.end
else
session("logintime")=rs("LastLoginTime") 
'conn.execute("update qyml set LastLoginIP='"&Request.ServerVariables("REMOTE_ADDR")&"',LastLoginTime='"&now&"',LoginTimes=LoginTimes+1 Where User='"&UsernameGet&"'") 
rs("LastLoginIP")=Request.ServerVariables("REMOTE_ADDR")
rs("LastLoginTime")=now()
rs("LoginTimes")=rs("LoginTimes")+1
rs.update
session("logintimes")=rs("LoginTimes")
session("user")=UsernameGet
session("id")=rs("id")
session("hydj")=rs("flag")
session("siji")=rs("siji")
session("uflag")=rs("uflag")
session("Qymc")=rs("Qymc")
session("name")=rs("name")
session("email")=rs("email")
session("address")=rs("address")
session("post")=rs("post")
session("Phone")=rs("Phone1")&"-"&rs("Phone2")&"-"&rs("Phone3")
session("Fax")=rs("Fax1")&"-"&rs("Fax2")&"-"&rs("Fax3")
session("LastLoginTime")=rs("LastLoginTime")
session.Timeout=30
		   if url="" then
           Response.Redirect ("../index.asp")
		   else
		    'response.Write "2秒后自动跳转到"&url
            response.write "<meta http-equiv=""refresh"" content=""1;URL="+url+""">"
           ' response.write 
		   end if

Response.End
rs.close 
set rs=nothing 
conn.close
set conn=nothing 
end if
%>
