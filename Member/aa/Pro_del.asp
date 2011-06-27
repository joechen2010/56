<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim id,fs,file2
id=request("id")
set rs=server.createobject("adodb.recordset")
sql="select * from spzs where id="&id
rs.open sql,conn,3,2
if not rs.eof then
file2=newspic&rs("Picture")
'response.write file2
set fs=server.CreateObject("scripting.filesystemobject")
file2=server.MapPath(file2)
if fs.FileExists(file2) then
fs.DeleteFile file2,true
end if
rs.delete
rs.update
rs.close
end if
response.redirect "Pro_Manage.asp"
%>