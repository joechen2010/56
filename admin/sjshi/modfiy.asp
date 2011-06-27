<!--#include file="../inc/syscode.asp" -->
<%
	  
Set rs = Server.CreateObject("ADODB.Recordset")    
	sele="select * from shejishi where id="&session("id")&""
	rs.Open sele,conn,3,2

	rs("sjname")=request("title")
	rs("linian")=request("linian")	
	rs("mome")=request("memo")	
	rs("e_mome")=request("e_memo")	
	rs("smallpic")=request("smallimg")		
	rs.update
	rs.close
	set rs=nothing
Response.Redirect"sjs_manage.asp"
%>