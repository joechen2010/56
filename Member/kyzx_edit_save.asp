<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim infoid,chufa,daoda,cartype,carload,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql ,i
infoid=request.form("infoid")
if isnumeric(infoid)=0 or infoid="" then
    response.write infoid
    response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	response.end
end if
chufa=trim(request.form("chufa"))
daoda=trim(request.form("daoda"))
cartype=trim(request("cartype"))
carload=request("carload")



set rs=server.createobject("adodb.recordset")
sql="select * from zx_info where infoid="&infoid 
rs.open sql,conn,1,3
if rs.eof and rs.bof then
response.write"<SCRIPT language=JavaScript>alert('对不起，此信息不存在！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs("province")=request("province")
rs("province2")=request("province2")
rs("city")=request("city")
rs("city2")=request("city2")
if request("area")<>"选择区县" then
rs("area")=request("area")
end if
if request("area2")<>"选择区县" then
rs("area2")=request("area2")
end if
'rs("name")=request("name")
'rs("tj_city")=request("tj_city")
'rs("carlong")=request("carlong")
'rs("validate")=request("validate")
rs("content")=request("content")
rs("gsid")=session("user")
rs("addtime")=now()
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('信息发布成功！');top.location.href='kyzx_Manage.asp';</script>"
	response.end
end if
%>