<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<%
Dim infoid,Action
infoid=Request.QueryString("id")
Action=trim(Request("Action"))
if isnumeric(infoid)=0 or infoid="" then
    response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	response.end
end if
set rs=server.createobject("adodb.recordset")
    sql="select * from zx_info2 where id="&infoid
    rs.open sql,conn,3,2
    if rs.eof and rs.bof then
    response.write "<script>alert('数据不存在！');top.location.href='hyzx_manage.asp';</script>"
    response.end
    end if
	rs.close
	set rs=nothing
if Action="Del" then
    conn.execute("delete from zx_info2 Where id="&infoid)
    response.write "<script>alert('数据删除成功！');top.location.href='zzz_add2.asp?infoid="&request("infoid")&"&city="&request("city")&"';</script>"
	
    response.end
elseif Action="Update" then
    conn.execute("update zx_info Set Addtime='"&Date()&"' Where ProID="&ProID)
	response.write "<script>alert('数据更新成功！');top.location.href='hyzx_manage.asp';</script>"
    response.end
else
    response.write "<script>alert('参数传递失败！');top.location.href='hyzx_manage.asp';</script>"
    response.end
end if
%>