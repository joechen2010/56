<%@ codepage ="936" %>
<%OPTION EXPLICIT%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<%
dim cpmc,cpbh,cpsb,cpcd,cpjg,cpgg,sortid,typeid,picture,jysm,xxsm
cpmc=request("cpmc")
cpbh=request("cpbh")
cpsb=request("cpsb")
cpcd=request("cpcd")
picture=request("picture")
jysm=request("jysm")
dim rs,sql
set rs=server.createobject("adodb.recordset")
sql="select * from spzs where cpmc='"&cpmc&"'and gsid="&gsid&"" 
rs.open sql,conn,1,3
if not rs.eof then
    response.write"<SCRIPT language=JavaScript>alert('对不起，您已经提交过此类产品！');"
    response.write"javascript:history.go(-1)</SCRIPT>"
    response.end 
else
    rs.addnew
    rs("cpbh")=cpbh &"-" & gsid
    rs("cpmc")=cpmc
    rs("cpsb")=cpsb
    rs("cpcd")=cpcd
    rs("picture")=picture
    rs("jysm")=dvHTMLEncode(jysm)
    rs("gsid")=gsid
    if session("flag")<>0 then
        rs("flag")=1
    else
        rs("flag")=0
    end if
    rs("idate")=date()
    rs.update
    rs.close
    set rs=nothing
    conn.close
    set conn=nothing
    response.redirect ("Pro_Manage.asp")
end if
%>
