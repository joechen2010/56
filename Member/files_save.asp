<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim carnum,carpp,cartype,Price,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql ,i
carnum=trim(request.form("carnum"))
carpp=trim(request.form("carpp"))
cartype=trim(request("cartype"))

set rs=server.createobject("adodb.recordset")
sql="select * from file_info where gsid='"&session("user")&"'and carnum='"&carnum&"' and carpp='"&carpp&"'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('对不起，您已经提交过此信息！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs.addnew
rs("infotype")="车辆档案"
rs("carnum")=request("carnum")
rs("carpp")=request("carpp")
rs("cartype")=request("cartype")
rs("carload")=request("carload")
rs("carlong")=request("carlong")
rs("time")=request("time")
rs("content")=request("content")
rs("siji")=request("siji")
rs("gsid")=session("user")
'rs("addtime")=date()
rs("addtime")=now()
rs("hits")=0
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('信息发布成功！');top.location.href='files_Manage.asp';</script>"
	response.end
end if
%>