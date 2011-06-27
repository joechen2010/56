<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim city2,cartype,Price,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql,i,city,infoid
city=trim(request.form("city"))
infoid=trim(request.form("infoid"))

set rs=server.createobject("adodb.recordset")
sql="select * from zx_info2 where city='"&request("city_tj")&"' and infoid="&request("infoid")&"" 
rs.open sql,conn,1,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('对不起，您已经提交过此信息！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
city=trim(request.form("city"))
rs.addnew
'rs("infotype")="货运专线"
rs("province")=request("province")
rs("city")=request("city_tj")
'rs("area")=request("area")
rs("infoid")=request("infoid")
'rs("tj_city")=request("tj_city")
'rs("validate")=request("validate")
rs("prices")=request("prices")
rs("tiji")=request("tiji")
'rs("gsid")=session("user")
'rs("addtime")=date()
rs("addtime")=now()
rs.update
'city=request("city")
infoid=rs("infoid")
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('信息发布成功！');top.location.href='zzz_add.asp?infoid="&infoid&"&city="&request("city")&"';</script>"

	response.end
end if
%>