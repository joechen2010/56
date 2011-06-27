<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<%
dim showname
dim ppxh
dim jgsm
dim mmfx
dim period
dim content
dim linkman
dim company
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
dim info_id
info_id=request("info_id")
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
email=request("email")
city=request.form("city")
set rs=server.createobject("adodb.recordset")
sql="select * from info where info_id="&info_id
rs.open sql,conn,1,3
rs("showname")=showname
rs("type")=mmfx
rs("period")=period
rs("content")=dvHTMLEncode(content)
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
rs("dateandtime")=date()
if session("flag")=1 then
   rs("flag")=1
else
   rs("flag")=0
end if
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('数据修改成功！');top.location.href='info_manage.asp';</script>"
response.end
%>