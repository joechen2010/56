<!--#include file = "../Inc/Syscode.asp"-->
<%
if request("action")="add" then
  set rs=server.CreateObject("adodb.recordset")
  sql="select * from advertise "
  rs.open sql,conn,1,3
  rs.addnew
  rs("adclass")=trim(request("adclass"))
  rs("webpic")=trim(request("bigimg"))
  rs("webname")=trim(request("webname"))
  rs("weburl")=trim(request("weburl"))
  rs("addtime")=now()
  rs.update
  rs.close
  set rs=nothing
  response.Write "<script>alert('��ӳɹ�');location.href='add.asp';</script>"
 elseif request("action")="modify" then 
  set rs=server.CreateObject("adodb.recordset")
  sql="select * from advertise where id="&request("id")&""
  rs.open sql,conn,1,3
  rs("adclass")=trim(request("adclass"))
  rs("webpic")=trim(request("bigimg"))
  rs("webname")=trim(request("webname"))
  rs("weburl")=trim(request("weburl"))
  rs("addtime")=now()
  rs.update
  rs.close
  set rs=nothing
  response.Write "<script>alert('�޸ĳɹ�');location.href='ad_manage.asp?page="&request("page")&"';</script>"  
end if

if request("manage")="" then 
    call add()
  elseif request("manage")="modify" then
    call modify()
end if
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../inc/Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" bgcolor="#e1eefd">
<p>&nbsp;</p>
<%sub add()%>
<table width="90%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#6699ff">
  <form method="POST" action="add.asp">
  <tr align="center"> 
    <td height="27" colspan="2" background="images/Left01.gif"><b>�� �� �� �� �� Ϣ</b></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="30" align="right">���ѡ��</td>
    <td height="30">
                          <select name="adclass">
                            <option value="1">�һ��ҳ�</option>
							<option value="2">����ר��</option>
							<option value="3">����ר��</option>
							<option value="4">��������</option>
							<option value="5">�������</option>
							<option value="6">�ⷿ��Ϣ</option>
							<option value="7">��Ƹ��ְ</option>
							<option value="8">������Ϣ</option>
                          </select>	</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">��˾���ƣ�</td>
      <td height="30" width="81%">
      <input name="webname" type="text" id="webname" style="border: 1px outset #C0C0C0" size="20">    </td>
  </tr>  
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">���ͼƬ��</td>
      <td height="30" width="81%">
 <iframe name="BigImg" frameborder=0 width=450 height=20 scrolling=no src=BigUpload.asp></iframe> 
            <br>
            ͼƬ·���� 
            <input name="BigImg" type="TEXT" id="BigImg"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100 readonly>  
	  </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="28" width="19%" align="right">������ַ��</td>
      <td width="81%" height="28">
	  <input name="weburl" type="text" id="weburl" style="border: 1px outset #C0C0C0" value="http://" size="20"></td>
  </tr>
    <tr bgcolor="#FFFFFF" align="center"> 
      <td height="30" colspan="2"> 
        <input type="submit" value=" �� �� "  style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="reset" value=" �� �� " name="B2" style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="hidden" name="action" value="add">      
	</td>
  </tr>
  </form>
</table>
<%end sub%>

<%sub modify()%>
<%
set web=conn.execute("select * from advertise where id="&request("id")&"")
if not (web.eof and web.bof) then

	adclass=web("adclass")
	select case adclass
	  case "1"
	  adtype="�һ��ҳ�"
	  case "2"
	  adtype="����ר��"
	  case "3"
	  adtype="����ר��"
	  case "4"
	  adtype="��������"
	  case "5"
	  adtype="�������"	  
	  case "6"
	  adtype="�ⷿ��Ϣ"
	  case "7"
	  adtype="��Ƹ��ְ"
	  case "8"
	  adtype="������Ϣ"	  	  
	end select
%>
<table width="90%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#6699ff">
  <form method="POST" action="add.asp">
  <tr align="center"> 
    <td height="27" colspan="2" background="images/Left01.gif"><b>�� �� �� �� �� Ϣ</b></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="30" align="right">���ѡ��</td>
    <td height="30">
                          <select name="adclass">
						    <option value="<%=web("adclass")%>"><%=adtype%></option>
                            <option value="1">�һ��ҳ�</option>
							<option value="2">����ר��</option>
							<option value="3">����ר��</option>
							<option value="4">��������</option>
							<option value="5">�������</option>
							<option value="6">�ⷿ��Ϣ</option>
							<option value="7">��Ƹ��ְ</option>
							<option value="8">������Ϣ</option>
                          </select>	</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">��˾���ƣ�</td>
      <td height="30" width="81%">
      <input name="webname" type="text" value="<%=web("webname")%>" style="border: 1px outset #C0C0C0" size="20">    </td>
  </tr>  
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">���ͼƬ��</td>
      <td height="30" width="81%">
 <iframe name="BigImg" frameborder=0 width=450 height=20 scrolling=no src=BigUpload.asp></iframe> 
            <br>
            ͼƬ·���� 
            <input name="BigImg" type="TEXT" id="BigImg" value="<%=web("webpic")%>"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100 readonly>  
	  </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="28" width="19%" align="right">������ַ��</td>
      <td width="81%" height="28">
	  <input name="weburl" type="text" id="weburl" value="<%=web("weburl")%>" style="border: 1px outset #C0C0C0" size="20"></td>
  </tr>
    <tr bgcolor="#FFFFFF" align="center"> 
      <td height="30" colspan="2"> 
    <input type="submit" value="�ύ"  style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="hidden" name="action" value="modify">
	<input type="hidden" name="id" value="<%=web("id")%>">
	<input type="hidden" name="page" value="<%=request("page")%>">      
	</td>
  </tr>
  </form>
</table>
<%
web.close
set web=nothing
end if
end sub
%>
</body>
</html>
