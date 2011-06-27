<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<%
Dim InfoID,Action
InfoID=Request.QueryString("InfoID")
Action=trim(Request("Action"))
if isnumeric(InfoID)=0 or InfoID="" then
    response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	response.end
end if
set rs=server.createobject("adodb.recordset")
    sql="select * from info where InfoID="&InfoID
    rs.open sql,conn,3,2
    if rs.eof and rs.bof then
    response.write "<script>alert('数据不存在！');top.location.href='info_manage.asp';</script>"
    response.end
    end if
	rs.close
	set rs=nothing
if Action="Del" then
    conn.execute("delete from Info Where InfoID="&InfoID)
    response.write "<script>alert('数据删除成功！');top.location.href='info_manage.asp';</script>"
    response.end
elseif Action="Update" then
    conn.execute("update Info Set Addtime='"&Date()&"' Where InfoID="&InfoID)
	response.write "<script>alert('数据更新成功！');top.location.href='info_manage.asp';</script>"
    response.end
else
    response.write "<script>alert('参数传递失败！');top.location.href='info_manage.asp';</script>"
    response.end
end if
%>