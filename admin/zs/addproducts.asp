<!--#include file="../inc/syscode.asp" -->
<%
Set rs = Server.CreateObject("ADODB.Recordset")    
	sele="select * from zs"
	rs.Open sele,conn,3,2
	rs.addnew
	rs("type")=request("type")
	rs("name")=request("name")
	rs("number")=request("number")		
	
	rs.update
	rs.close
	set rs=nothing
Response.Redirect"add.asp"
%>