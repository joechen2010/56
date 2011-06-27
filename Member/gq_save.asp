<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim title,cartype,Price,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql ,i
title=trim(request.form("title"))
infotype=trim(request.form("infotype"))
kftype=trim(request("kftype"))

set rs=server.createobject("adodb.recordset")
sql="select * from gq_info where gsid='"&session("user")&"'and title='"&title&"' and infotype='"&infotype&"'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('对不起，您已经提交过此信息！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs.addnew
rs("title")=request("title")
rs("infotype")=request("infotype")
rs("province")=request("province")
rs("city")=request("city")
if request("area")<>"选择区县" then
rs("area")=request("area")
end if
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
response.write "<script>alert('信息发布成功！');top.location.href='gq_Manage.asp';</script>"
	response.end
end if
%>