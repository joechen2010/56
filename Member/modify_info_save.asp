<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim rs,sql
dim user,pass
dim question
dim answer
dim sf
dim city
dim post
dim address
dim phone1,phone2,phone3
dim mobile
dim fax1,fax2,fax3
dim email
dim web
dim name
dim ch
dim bm
dim zw 
dim sortid
dim typeid
dim errmsg
question=request.form("question")
answer=request.form("answer")
province=request.form("province")
mobile=request.form("mobile")
city=request.form("city")
area=request.form("area")
post=request.form("post")
address=request.form("address")
phone1=request.form("phone1")
phone2=request.form("phone2")
phone3=request.form("phone3")
email=request.form("email")
name=request.form("name")
ch=request.form("ch")
bm=request.form("bm")
zw=request.form("zw")
mobile=request.form("mobile")
zw=request.form("zw")
fax1=request.form("fax1")
fax2=request.form("fax2")
fax3=request.form("fax3")
web=request.form("web")
if web="http://" then web=""
if len(web)>7 and left(web,7) <> "http://" then web="http://" & web
%>
<%
set rs=server.CreateObject ("adodb.recordset")
sql="select * from qyml where User='"&session("user")&"'"
rs.open sql,conn,1,3
rs("question")=question
rs("answer")=answer
rs("province")=province
rs("city")=city
rs("area")=area
rs("post")=post
rs("address")=address
rs("mobile")=mobile
rs("phone1")=phone1
rs("phone2")=phone2
rs("phone3")=phone3
rs("fax1")=fax1
rs("fax2")=fax2
rs("fax3")=fax3
rs("email")=email
rs("web")=web
rs("name")=name
rs("ch")=ch
rs("bm")=bm
rs("zw")=zw
rs("gsid")=user
rs.update
session("name")=rs("name")
session("email")=rs("email")
session("address")=rs("address")
session("post")=rs("post")
session("Phone")=rs("Phone1")&"-"&rs("Phone2")&"-"&rs("Phone3")
session("Fax")=rs("Fax1")&"-"&rs("Fax2")&"-"&rs("Fax3")
rs.close
conn.close
set conn=nothing 
Response.Redirect "modify_info.asp"
%>
