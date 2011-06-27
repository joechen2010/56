<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim showname
dim mmfx
dim period
dim content
dim linkman
dim address
dim email
dim phone
dim fax
dim postcode
dim web
dim city
dim country
dim rs
dim sql 
showname=request("showname")
mmfx=request("type")
period=request("period")
content=request("content")
linkman=request("linkman")
company=request("company")
address=request("address")
email=request("email")
phone=request("phone")
country=request("country")
post=request("postcode")
fax=request.form("fax")
web=request.form("web")
if web="http://" then web=""
if len(web)>7 and left(web,7) <> "http://" then web="http://" & web
city=request.form("city")
set rs=server.createobject("adodb.recordset")
sql="select * from info where company='"&request("company")&"'and showname='"&request("showname")&"'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('对不起，您已经提交过此信息！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs.addnew
rs("showname")=showname
rs("type")=mmfx
rs("period")=period
rs("content")=content
rs("linkman")=linkman
rs("company")=company
rs("address")=address
rs("mail")=email
rs("phone")=phone
rs("fax")=fax
rs("postcode")=post
rs("country")=country
rs("web")=web
rs("city")=city
rs("gsid")=session("id")
rs("dateandtime")=date()
rs("tj")=0
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('商机添加成功！');top.location.href='Info_Manage.asp';</script>"
	response.end
end if
%>