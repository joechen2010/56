<!--#include file="../inc/conn.asp"-->
<%
sql="select * from Qyzs where id="+session("id")+""
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,3,3 
if not rs.eof then
rs("zsmc")=request.form("zsmc")
rs("zsurl")=request("smallimg")
rs.update
response.Write "<script language='javascript'>alert('�޸ĳɹ���');window.location.href='qyry.asp';</script>"
else
response.Write "<script language='javascript'>alert('�˼�¼�����ڣ�');window.location.href='qyry.asp';</script>"
end if
%>


