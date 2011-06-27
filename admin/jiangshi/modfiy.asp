<!--#include file="../inc/syscode.asp" -->
<%
	  
Set rs = Server.CreateObject("ADODB.Recordset")    
	sele="select * from jsfc where id="&session("id")&""
	rs.Open sele,conn,3,2

	rs("jsname")=request("title")
	rs("smallpic")=request("smallimg")	
	rs("zhiwu")=request("zhiwu")
	rs("zhicheng")=request("zhicheng")
	rs("age")=request("age")	
		rs("jiang")=request("linian")	
		rs("produ")=request("e_memo")	
	rs.update
	rs.close
	set rs=nothing
Response.Redirect"sjs_manage.asp"
%>