<%@ codepage ="936" %>
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<%
dim cpbh,cpmc,cpsb,cpcd,cpjg,cpgg,bighy,lithy,jysm,xxsm
id=request("id")
cpbh=request("cpbh")
cpmc=request("cpmc")
cpsb=request("cpsb")
cpcd=request("cpcd")
cpjg=request("cpjg")
picture=request("picture")
cpgg=request("cpgg")
sortid=request("sortid")
typeid=request("typeid")
jysm=request("jysm")
xxsm=request("xxsm")
gsid=session("id")
If request("zsbz")="" Then
zsbz=0
else
zsbz=1
end if

set rs=server.createobject("adodb.recordset")
sql="select * from spzs where id="&id
rs.open sql,conn,1,3
rs("cpbh")=cpbh
rs("cpmc")=cpmc
rs("cpsb")=cpsb
rs("cpcd")=cpcd
rs("cpjg")=cpjg
rs("cpgg")=cpgg
rs("picture")=picture
rs("sortid")=sortid
rs("typeid")=typeid
rs("jysm")=dvHTMLEncode(jysm)
rs("xxsm")=dvHTMLEncode(xxsm)
rs("idate")=date()
rs.update
rs.close
set rs=nothing
response.write "<script>alert('产品修改成功！');top.location.href='Pro_Manage.asp';</script>"
	response.end
%>