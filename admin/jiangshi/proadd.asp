<!--#include file="../inc/syscode.asp" -->
<%
	  
Set rss = Server.CreateObject("ADODB.Recordset")    
	seles="select * from shejishi where id="&request("id")&""
	rss.Open seles,conn,3,2
	  
Set rs= Server.CreateObject("ADODB.Recordset")    
	sele="select * from pic"
	rs.Open sele,conn,3,3
	rs.addnew
	rs("tid")=request("id")
	rs("smallpic")=request("smallimg")	
	rs("bigpic")=request("bigimg")	
	rs("sjname")=rss("sjname")	
	rs.update
	rs.close
	set rss=nothing
Response.Redirect"product.asp"
%>