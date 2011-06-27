<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../Member/check.asp"-->
<%
Dim infoid,Action
infoid=Request.QueryString("infoid")
Action=trim(Request("Action"))
if isnumeric(ProID)=0 or infoid="" then
    response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	response.end
end if
set rs=server.createobject("adodb.recordset")
    sql="select * from gq_info where infoid="&infoid
    rs.open sql,conn,3,2
    if rs.eof and rs.bof then
     response.Redirect "gq_manage.asp"
    response.end
    end if
	rs.close
	set rs=nothing
if Action="Del" then
    conn.execute("delete from gq_info Where infoid="&infoid)
    response.Redirect "gq_manage.asp"
    response.end
elseif Action="Update" then
    conn.execute("update gq_info Set Addtime='"&Date()&"' Where ProID="&ProID)
	response.write "<script>alert('数据更新成功！');top.location.href='gq_manage.asp';</script>"
    response.end
else
    response.write "<script>alert('参数传递失败！');top.location.href='gq_manage.asp';</script>"
    response.end
end if
%>