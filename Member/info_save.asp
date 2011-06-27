<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim title,infotype,keyword,validate,SortID,TypeID,sContent,PicImg

dim sql ,i
title=trim(request.form("title"))
infotype=request.form("infotype")
keyword=trim(request("keyword"))
validate=request("validate")
SortID=request("SortID")
TypeID=request("TypeID")
'sContent=""
'For i=1 to Request.Form("content").count
   'sContent=sContent & Request.Form("content")(i)
'Next
ProImg=request("ProImg")
content=request("content")
content=replace(content,vbcrlf,"<br>")
        content=replace(content," ","&nbsp;")
set rs=server.createobject("adodb.recordset")
sql="select * from info" 
rs.open sql,conn,1,3
rs.addnew
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
if session("hydj")=0 then
   rs("flag")=0
else
   rs("flag")=1
end if
rs("hits")=0
rs("bz")=0
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('信息发布成功！');top.location.href='Info_Manage.asp';</script>"
response.end
%>