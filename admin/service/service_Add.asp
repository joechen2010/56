<!--#include file = "../Inc/Syscode.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Admin_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

set rs=server.CreateObject("adodb.recordset")
sql="select top 1 * from service"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
  rs.addnew
  rs.update
end if
action=request("action")
select case action
   case "about"
   title="��������"
   content=server.HTMLEncode(rs("about")&"&nbsp;")
   case "service"
   title="��������"
   content=server.HTMLEncode(rs("service")&"&nbsp;")
   case "guide"
   title="��վָ��"
   if not isnull(rs("guide")) then 
   content=server.HTMLEncode(rs("guide"))
   end if
   case "charges"
   title="�ʷѱ�׼"
   content=server.HTMLEncode(rs("charges")&"&nbsp;")
   case "ads"
   title="��վ���"
   content=server.HTMLEncode(rs("ads")&"&nbsp;")
   case "contact"
   title="��ϵ����"
   content=server.HTMLEncode(rs("contact")&"&nbsp;")
end select   
%>
  <form action="service_add.asp" method="post" name="myform">
	
  <table width="650" align=center cellspacing=3>
    <tr> 
      <td colspan="2" align="center"></td>
    </tr>
    <tr> 
      <td width="71" align="right"><%=title%>:</td>
      <td width="564"> <input name="Content" type="hidden" value="<%=content%>">
        <iframe ID="eWebEditor1" src="../Editor/ewebeditor.asp?id=Content&style=" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>      </td>
    </tr>
  </table>
	<p align=center><input type=submit name=btnSubmit value=" �� �� "> <input type=reset name=btnReset value=" �� �� "></p>
	<input type="hidden" name="action2" value="<%=action%>"><input type="hidden" name="id" value="<%=rs("id")%>"
</form>
<%
rs.close
set rs=nothing
%>
</body>
</html>
<%
if request("action2")<>"" then
   action=request("action2")
   id=request("id")
   content=request("content")
  call modify(action,content,id)
end if
sub modify(action,content,id)
  conn.execute("update service set "&action&"='"&content&"' where id="&id)
  response.Write "<script>alert('�޸ĳɹ�');location.href='service_add.asp?action="&action&"';</script>"
  'response.Redirect "service_add.asp?action="&action&""
end sub

%>



