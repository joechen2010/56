<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim title,infotype,keyword,validate,SortID,TypeID,sContent,PicImg

dim sql ,i,infoid
infoid=request.form("infoid")
if isnumeric(infoid)=0 or infoid="" then
    response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	response.end
end if
title=trim(request.form("title"))
infotype=request.form("infotype")
keyword=trim(request("keyword"))
validate=request("validate")
SortID=request("SortID")
TypeID=request("TypeID")
sContent=""
'For i=1 to Request.Form("content").count
   'sContent=sContent & Request.Form("content")(i)
'Next
ProImg=request("ProImg")
content=request("content")
content=replace(content,vbcrlf,"<br>")
content=replace(content," ","&nbsp;")
set rs=server.createobject("adodb.recordset")
sql="select * from info where infoid="&infoid 
rs.open sql,conn,1,3
if rs.eof and rs.bof then
response.write"<SCRIPT language=JavaScript>alert('�Բ��𣬴���Ϣ�����ڣ�');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs("title")=title
rs("infotype")=infotype
rs("keyword")=keyword
rs("validate")=validate
rs("SortID")=SortID
rs("TypeID")=TypeID
rs("content")=content
rs("ProImg")=ProImg
rs("gsid")=session("user")
'rs("addtime")=date()
rs("addtime")=now()
rs("bz")=0
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('��Ϣ�����ɹ���');top.location.href='Info_Manage.asp';</script>"
	response.end
end if
%>