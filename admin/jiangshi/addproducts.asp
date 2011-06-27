<!--#include file="../inc/syscode.asp" -->
<%

	    body1=request("e_memo")
		body1=replace(body1,vbcrlf,"<br>")
       body1=replace(body1," ","&nbsp;")	
	   
	   body2=request("linian")
		body2=replace(body2,vbcrlf,"<br>")
       body2=replace(body2," ","&nbsp;")	  
Set rs = Server.CreateObject("ADODB.Recordset")    
	sele="select * from jsfc"
	rs.Open sele,conn,3,2

	'ss=obj.SaveFile("small",Server.MapPath("smallpic"), true) '保存文件到服务器	
	'bb=obj.SaveFile("big",Server.MapPath("pic"), true) '保存文件到服务器	
	rs.addnew
	rs("jsname")=request("title")
	rs("jiang")=body2	
	rs("zhiwu")=request("zhiwu")
	rs("zhicheng")=request("zhicheng")
	rs("age")=request("age")
	rs("produ")=body1	
	

	rs("smallpic")=request("smallimg")		
	
	rs.update
	rs.close
	set rs=nothing
Response.Redirect"add.asp"
%>