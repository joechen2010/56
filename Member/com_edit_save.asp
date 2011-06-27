<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim rs,sql,qymc,zycp,qyxz,nyye,ygrs,frdb,zczj,sortid,qylb,types,jygs,sqyjj
dim errmsg
qymc=request.form("qymc")
zycp=request.form("zycp")
qyxz=request.form("qyxz")
nyye=request.form("nyye")
ygrs=request.form("ygrs")
frdb=request.form("frdb")
if request("zczj")="0"then
zczj="无需验资"
else
zczj=request.form("zczj")
end if
sortid=request.form("sortid")
qylb=request.form("qylb")
if sortid="1" then
    types=request.form("fullbearings1")
elseif sortid="2" then
    types=request.form("fullbearings2")
elseif sortid="3" then
    types=request.form("fullbearings3")
else
    types=""
end if
jygs=request.form("jygs")
'sqyjj = ""
'For i = 1 To Request.Form("qyjj").Count
	'sqyjj = sqyjj & Request.Form("qyjj")(i)
'Next
content=request("content")
content=replace(content,vbcrlf,"<br>")
content=replace(content," ","&nbsp;")

set rs=server.CreateObject ("adodb.recordset")
sql="select * from qyml where user='"&session("user")&"'"
rs.open sql,conn,1,3
		  rs("siji")=request("siji")
          rs("qylb")=request("SortID")
		  if trim(request("Province"))<>"省份" then
          rs("Province")=request("Province")
		  end if
		  if trim(request("city"))<>"地级市" then
          rs("City")=request("city")
		  end if
		  if trim(request("area"))<>"选择区县" then
		  rs("Area")=request("area")
		  end if
          rs("address")=request("address")
          rs("post")=request("post")
          'rs("zw")=zw
          rs("phone1")=request("phone1")
		  rs("phone2")=request("phone2")
		  rs("phone3")=request("phone3")
          rs("mobile")=request("mobile")
          rs("fax1")=request("fax1")
		  rs("fax2")=request("fax2")
		  rs("fax3")=request("fax3")
          rs("email")=request("email")
          rs("qymc")=request("qymc")
          rs("web")=request("web")
          rs("qyjj")=dvHTMLCode(request("content"))
rs.update
session("Qymc")=rs("Qymc")
rs.close
set rs=nothing 
conn.close
set conn=nothing 
response.write "<SCRIPT language=JavaScript>alert('您的资料已经发布成功！');"
response.write"this.location.href='index.asp';</SCRIPT>" 
%>