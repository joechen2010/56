<!--#include file="../inc/conn.asp"-->
<%
sql="select * from Qyzs where id="+session("id")+""
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,3,3 
if not rs.eof then
rs("zsmc")=request.form("zsmc")
rs("zsurl")=request("smallimg")
rs.update
response.Write "<script language='javascript'>alert('修改成功！');window.location.href='qyry.asp';</script>"
else
response.Write "<script language='javascript'>alert('此记录不存在！');window.location.href='qyry.asp';</script>"
end if
%>


