<%@ codepage ="936" %>
<!--#include file="../inc/CONN.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim BearinNo,Brand,Num,Price,SortID,TypeID,Nowplace,MadePlace,sIntroduce,PicImg

dim sql ,i
BearinNo=trim(request.form("BearinNo"))
Brand=request.form("Brand")
Num=trim(request("Num"))
Price=request("Price")
SortID=request("SortID")
TypeID=request("TypeID")
Nowplace=request("Nowplace")
MadePlace=request("MadePlace")
sIntroduce=""
For i=1 to Request.Form("Introduce").count
   sIntroduce=sIntroduce & Request.Form("Introduce")(i)
Next
ProImg=request("ProImg")

set rs=server.createobject("adodb.recordset")
sql="select * from spzs where gsid='"&session("user")&"'and BearinNo='"&BearinNo&"'" 
rs.open sql,conn,1,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('对不起，您已经提交过此信息！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end 
else
rs.addnew
rs("BearinNo")=BearinNo
rs("Brand")=Brand
rs("Num")=Num
rs("Price")=Price
rs("SortID")=SortID
rs("TypeID")=TypeID
rs("Nowplace")=Nowplace
rs("MadePlace")=MadePlace
rs("Introduce")=sIntroduce
rs("ProImg")=ProImg
rs("gsid")=session("user")
rs("addtime")=date()
if session("hydj")=0 then
   rs("flag")=0
else
   rs("flag")=1
end if
rs("hits")=0
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script>alert('信息发布成功！');top.location.href='Pro_Manage.asp';</script>"
	response.end
end if
%>