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
    response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	response.end
end if

set rs=server.createobject("adodb.recordset")
sql="select * from zp_info where infoid="&infoid 
rs.open sql,conn,1,3
if rs.eof and rs.bof then
response.write"<SCRIPT language=JavaScript>alert('�Բ��𣬴���Ϣ�����ڣ�');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs("province")=request("province")
rs("city")=request("city")
rs("area")=request("area")
rs("job")=request("job")
rs("money")=request("money")
rs("jobtype")=request("jobtype")
rs("contact")=request("contact")
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
response.write "<script>alert('��Ϣ�����ɹ���');top.location.href='zp_Manage.asp';</script>"
	response.end
end if
%>