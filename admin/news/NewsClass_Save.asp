<!--#include file = "../Inc/Syscode.asp"-->
<%

dim ClassID,rs,sql,sqltext
dim WriteErrMsg
if request("action")="add" then 
call addNsclass()
elseif request("action")="modify" then
call modifyNsClass()
end if

'��Ӽ�¼
sub addNsclass()
set rs=server.createobject("adodb.recordset")
sqltext="select * from News_Class where ClassName='"&request("ClassName")&"' "
rs.open sqltext,conn,1,3
if rs.recordcount >= 1 then 
   response.Write "<script language='javascript'>alert('����Ŀ�Դ��ڣ�');history.back();</script>"
else
   rs.addnew
   rs("ClassName")=trim(request.form("ClassName"))
   rs.update
   rs.close
   Response.Write "<script language=javascript>alert('������ӳɹ���');location.href='NewsClass_Manage.asp';</script>"
end if
end sub

'�޸ļ�¼

sub modifyNsClass()
ClassID=request("ClassID")
set rs=server.createobject("adodb.recordset")
sqltext="select * from News_Class where ClassID="&ClassID
rs.open sqltext,conn,1,3
   rs("ClassName")=trim(request.form("ClassName"))
   rs.update
   rs.close
   Response.Write "<script language=javascript>alert('�����޸ĳɹ���');location.href='NewsClass_Manage.asp';</script>"
end sub

%>