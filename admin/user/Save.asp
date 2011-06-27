<%@ codepage ="936" %>
<!--#include file="../../Inc/Conn.asp"--> <link rel="stylesheet" type="text/css" href="../style.css"> 
<%
set rs=server.CreateObject ("adodb.recordset")
sql="select * from qyml where id="&request("id")&""
rs.open sql,conn,1,3
rs("user")=request.form("user")
rs("pass")=request.form("pass")
rs("question")=request.form("question")
rs("answer")=request.form("answer")
rs("qymc")=request.form("qymc")
rs("qyxz")=request.form("qyxz")
rs("zczj")=request.form("zczj")
rs("frdb")=request.form("frdb")
rs("ygrs")=request.form("ygrs")
rs("zh")=request.form("zh")
rs("bank")=request.form("bank")
rs("nyye")=request.form("nyye")
rs("zycp")=request.form("zycp")
rs("qyjj")=request.form("qyjj")
rs("sf")=request.form("sf")
rs("city")=request.form("city")
rs("post")=request.form("post")
rs("address")=request.form("address")
rs("phone")=request.form("phone")
rs("fax")=request.form("fax")
rs("email")=request.form("email")
rs("web")=request.form("web")
rs("name")=request.form("name")
rs("ch")=request.form("ch")
rs("bm")=request.form("bm")
rs("zw")=request.form("zw")
rs("gsid")=request.form("user")
rs("flag")=request("flag")
rs.update
rs.close
set rs=noting
conn.close
set conn=nothing
response.write"<SCRIPT language=JavaScript> this.location.href='javascript:window.history.go(-2)';</SCRIPT>" 
%>