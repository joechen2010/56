<%@language=vbscript codepage=936 %>
<%
option explicit

'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
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
	 call showerror("�û��������벻����!")
else
	 if password<>rs("password") then
	     call CloseConn
		 call showerror("�û������������") 
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