<%@language=vbscript codepage=936 %>
<%
option explicit

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%> 
<!--#include file = "../Inc/conn.asp"-->
<!--#include file = "../Inc/function.asp"-->
<!--#include file = "../Inc/md5.asp"-->
<%
dim sql,rs
dim username,password
username=replace(trim(request("username")),"'","")
password=replace(trim(Request("password")),"'","")
password=md5(password)
set rs=Server.CreateObject("Adodb.RecordSet")
    sql="select * from admin where password='"&password&"' and username='"&username&"'"
	rs.open sql,conn,1,1
	
if rs.bof and rs.eof then
	 call showerror("用户名或密码不存在!")
else
	 if password<>rs("password") then
	     call CloseConn
		 call showerror("用户名或密码错误！") 
	 else
	     conn.execute("update Admin set LastLoginIP='"&Request.ServerVariables("REMOTE_ADDR")&"',LastLoginTime='"&now&"',LoginTimes=LoginTimes+1 Where username='"&username&"'") 
		 session("AdminName")=username
                 session("AdminPower")=rs("purview")
		 rs.close
		 set rs=nothing
		 call CloseConn
		 response.Redirect "../Menu/Admin_index.asp"
	 end if
end if
%>