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
sql="select * from zx_info where gsid='"&session("user")&"'and city='"&city&"' and city2='"&city2&"' and infotype='����ר��'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('�Բ������Ѿ��ύ������Ϣ��');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs.addnew
rs("infotype")="����ר��"
rs("province")=request("province")
rs("province2")=request("province2")
rs("city")=request("city")
rs("city2")=request("city2")
if request("area")<>"ѡ������" then
rs("area")=request("area")
end if
if request("area2")<>"ѡ������" then
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
response.write "<script>alert('��Ϣ�����ɹ���');top.location.href='hyzx_Manage.asp';</script>"
	response.end
end if
%>