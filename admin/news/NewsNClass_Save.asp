<!--#include file = "../Inc/Syscode.asp"-->
<%
dim  ClassID,NClassName,rs,sql,sqltext
ClassID=trim(request("ClassID"))
NClassName=trim(request("NClassName"))
if request("action")="add" then
    call addNsNClass()
elseif request("action")="modify" then
    call modifyNsNClass()
end if

sub addNsNClass()
set rs=server.createobject("adodb.recordset")
sqltext="select * from News_NClass where NClassName='"&NClassName&"' and ClassID="&ClassID&""
rs.open sqltext,conn,1,3
if rs.recordcount >= 1 then 
    response.Write "<script language='javascript'>alert('该栏目以存在！');history.back();</script>"
else
    rs.addnew
    rs("NClassName")=NClassName
    rs("ClassID")=ClassID
    rs.update
    rs.close
    Response.Write "<script language=javascript>alert('数据添加成功！');location.href='NewsClass_Manage.asp';</script>"
end if
end sub


sub modifyNsNClass()
set rs=server.createobject("adodb.recordset")
sqltext="select * from News_NClass where NClassID="&request("NClassID")
rs.open sqltext,conn,1,3
rs("NClassName")=NClassName
rs("ClassID")=ClassID
rs.update
rs.close
   Response.Write "<script language=javascript>alert('数据修改成功！');location.href='NewsClass_Manage.asp';</script>"
end sub

%>