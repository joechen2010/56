<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim infoid,chufa,daoda,cartype,carload,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql ,i
infoid=request.form("id")
if isnumeric(infoid)=0 or infoid="" then
    response.write infoid
    response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	response.end
end if
chufa=trim(request.form("chufa"))
daoda=trim(request.form("daoda"))
cartype=trim(request("cartype"))
carload=request("carload")



set rs=server.createobject("adodb.recordset")
sql="select * from zx_info2 where id="&infoid 
rs.open sql,conn,1,3
if rs.eof and rs.bof then
response.write"<SCRIPT language=JavaScript>alert('�Բ��𣬴���Ϣ�����ڣ�');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
'rs("infotype")="����ר��"
rs("province")=request("province")
rs("city")=request("city_tj")
'rs("area")=request("area")
'rs("infoid")=request("infoid")
'rs("tj_city")=request("tj_city")
'rs("validate")=request("validate")
rs("prices")=request("prices")
rs("tiji")=request("tiji")
'rs("gsid")=session("user")
'rs("addtime")=date()
rs("addtime")=now()
rs.update
rs.close
	set rs=nothing
conn.close
set conn=nothing
end if
response.write "<script>alert('��Ϣ�����ɹ���');top.location.href='zzz_add.asp?infoid="&request("infoid")&"&city="&request("city")&"';</script>"
	response.end

%>