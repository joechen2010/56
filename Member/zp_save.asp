<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim city,city2,cartype,Price,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql ,i
city=trim(request.form("city"))
job=trim(request.form("job"))
money=trim(request.Form("money"))
jobtype=trim(request("jobtype"))
if request("action")="��Ƹ" then
set rs=server.createobject("adodb.recordset")
sql="select * from zp_info where gsid='"&session("user")&"'and city='"&city&"' and job='"&job&"' and jobtype='"&jobtype&"'and infotype='��Ƹ��Ϣ'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('�Բ������Ѿ��ύ������Ϣ��');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs.addnew
rs("infotype")="��Ƹ��Ϣ"
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
 else
  set rs=server.createobject("adodb.recordset")
sql="select * from zp_info where gsid='"&session("user")&"'and city='"&city&"' and job='"&job&"' and jobtype='"&jobtype&"' and infotype='��ְ��Ϣ'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('�Բ������Ѿ��ύ������Ϣ��');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs.addnew
rs("infotype")="��ְ��Ϣ"
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
rs("addtime")=date()
rs("hits")=0
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('��Ϣ�����ɹ���');top.location.href='zp_Manage.asp';</script>"
	response.end
end if
end if
%>