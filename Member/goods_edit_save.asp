<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim infoid,city,city2,cartype,carload,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql ,i
infoid=request.form("infoid")
if isnumeric(infoid)=0 or infoid="" then
    response.write infoid
    response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	response.end
end if
city=trim(request.form("city"))
city2=trim(request.form("city2"))
cartype=trim(request("cartype"))
carload=request("carload")



set rs=server.createobject("adodb.recordset")
sql="select * from info2 where infoid="&infoid 
rs.open sql,conn,1,3
if rs.eof and rs.bof then
response.write"<SCRIPT language=JavaScript>alert('�Բ��𣬴���Ϣ�����ڣ�');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
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
rs("cartype")=request("cartype")
rs("carload")=request("carload")
rs("carlong")=request("carlong")
rs("validate")=request("validate")
rs("content")=request("content")
rs("gsid")=session("user")
'rs("addtime")=date()
rs("addtime")=now()
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('��Ϣ�����ɹ���');top.location.href='goods_Manage.asp';</script>"
	response.end
end if
%>