<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<%
Dim infoid,Action
infoid=Request.QueryString("infoid")
Action=trim(Request("Action"))
if isnumeric(ProID)=0 or infoid="" then
    response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	response.end
end if
set rs=server.createobject("adodb.recordset")
    sql="select * from zx_info where infoid="&infoid
    rs.open sql,conn,3,2
    if rs.eof and rs.bof then
    response.write "<script>alert('���ݲ����ڣ�');top.location.href='kyzx_manage.asp';</script>"
    response.end
    end if
	rs.close
	set rs=nothing
if Action="Del" then
    conn.execute("delete from zx_info Where infoid="&infoid)
    response.write "<script>alert('����ɾ���ɹ���');top.location.href='kyzx_manage.asp';</script>"
    response.end
elseif Action="Update" then
    conn.execute("update zx_info Set Addtime='"&Date()&"' Where ProID="&ProID)
	response.write "<script>alert('���ݸ��³ɹ���');top.location.href='kyzx_manage.asp';</script>"
    response.end
else
    response.write "<script>alert('��������ʧ�ܣ�');top.location.href='kyzx_manage.asp';</script>"
    response.end
end if
%>