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
    sql="select * from info2 where infoid="&infoid
    rs.open sql,conn,3,2
    if rs.eof and rs.bof then
response.Redirect "car_manage.asp"
    response.end
    end if
	rs.close
	set rs=nothing
if Action="Del" then
    conn.execute("delete from info2 Where infoid="&infoid)
	response.Redirect "car_manage.asp"
    response.end
elseif Action="Update" then
    conn.execute("update info2 Set Addtime='"&Date()&"' Where ProID="&ProID)
	response.Redirect "car_manage.asp"
    response.end
else
    response.write "<script>alert('参数传递失败！');top.location.href='car_manage.asp';</script>"
    response.end
end if
%>