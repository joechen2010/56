<!--#include file="../inc/syscode.asp" -->
<%
Set rs = Server.CreateObject("ADODB.Recordset")    
	sele="select * from web"
	rs.Open sele,conn,3,2
	rs.addnew
	rs("type")=request("type")
	rs("webname")=request("webname")
	rs("weburl")=request("weburl")		
	
	rs.update
	rs.close
	set rs=nothing
Response.Redirect"add.asp"
%>