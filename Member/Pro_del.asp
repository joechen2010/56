<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<%
Dim ProID,Action
ProID=Request.QueryString("ProID")
Action=trim(Request("Action"))
if isnumeric(ProID)=0 or ProID="" then
    response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	response.end
end if
set rs=server.createobject("adodb.recordset")
    sql="select * from spzs where ProID="&ProID
    rs.open sql,conn,3,2
    if rs.eof and rs.bof then
    response.write "<script>alert('���ݲ����ڣ�');top.location.href='Pro_manage.asp';</script>"
    response.end
    end if
	rs.close
	set rs=nothing
if Action="Del" then
    conn.execute("delete from spzs Where ProID="&ProID)
    response.write "<script>alert('����ɾ���ɹ���');top.location.href='Pro_manage.asp';</script>"
    response.end
elseif Action="Update" then
    conn.execute("update spzs Set Addtime='"&Date()&"' Where ProID="&ProID)
	response.write "<script>alert('���ݸ��³ɹ���');top.location.href='Pro_manage.asp';</script>"
    response.end
else
    response.write "<script>alert('��������ʧ�ܣ�');top.location.href='Pro_manage.asp';</script>"
    response.end
end if
%>