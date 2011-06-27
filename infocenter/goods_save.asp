<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../Member/check.asp"-->
<%
dim chufa,daoda,cartype,Price,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql ,i
city=trim(request.form("city"))
city2=trim(request.form("city2"))
cartype=trim(request("cartype"))

set rs=server.createobject("adodb.recordset")
sql="select * from info2 where gsid='"&session("user")&"'and city='"&city&"' and city2='"&city2&"' and infotype='货源信息'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('对不起，您已经提交过此信息！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs.addnew
rs("infotype")="货源信息"
rs("province")=request("province")
rs("province2")=request("province2")
rs("city")=request("city")
rs("city2")=request("city2")
rs("area")=request("area")
rs("area2")=request("area2")
rs("cartype")=request("cartype")
rs("carload")=request("carload")
rs("carlong")=request("carlong")
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
response.write "<script>alert('信息发布成功！');</script>"
response.Redirect "car_manage.asp"
	response.end
end if
%>