<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<%
Dim InfoID,Action
InfoID=Request.QueryString("InfoID")
Action=trim(Request("Action"))
if isnumeric(InfoID)=0 or InfoID="" then
    response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	response.end
end if
set rs=server.createobject("adodb.recordset")
    sql="select * from info where InfoID="&InfoID
    rs.open sql,conn,3,2
    if rs.eof and rs.bof then
    response.write "<script>alert('���ݲ����ڣ�');top.location.href='info_manage.asp';</script>"
    response.end
    end if
	rs.close
	set rs=nothing
if Action="Del" then
    conn.execute("delete from Info Where InfoID="&InfoID)
    response.write "<script>alert('����ɾ���ɹ���');top.location.href='info_manage.asp';</script>"
    response.end
elseif Action="Update" then
    conn.execute("update Info Set Addtime='"&Date()&"' Where InfoID="&InfoID)
	response.write "<script>alert('���ݸ��³ɹ���');top.location.href='info_manage.asp';</script>"
    response.end
else
    response.write "<script>alert('��������ʧ�ܣ�');top.location.href='info_manage.asp';</script>"
    response.end
end if
%>