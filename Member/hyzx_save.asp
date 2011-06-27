<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim city,city2,cartype,Price,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql ,i
city=trim(request.form("city"))
city2=trim(request.form("city2"))
cartype=trim(request("cartype"))

set rs=server.createobject("adodb.recordset")
sql="select * from zx_info where gsid='"&session("user")&"'and city='"&city&"' and city2='"&city2&"' and infotype='货运专线'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('对不起，您已经提交过此信息！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs.addnew
rs("infotype")="货运专线"
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
rs("name")=request("name")
rs("tj_city")=request("tj_city")
rs("validate")=request("validate")
rs("content")=request("content")
rs("gsid")=session("user")
'rs("addtime")=date()
rs("addtime")=now()
rs("hits")=0
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('信息发布成功！');top.location.href='hyzx_Manage.asp';</script>"
	response.end
end if
%>