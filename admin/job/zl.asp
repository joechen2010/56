
<!--#include file = "../Inc/Syscode.asp"-->
<%

uid_=request("uid")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from geren where id="&uid_
rs.open sql,conn,1,1
if rs.eof then
response.write "��������"
response.end
end if
%>
<html>
<head>
<title><%=rs("iname")%>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="images/Admin.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table border="0" cellpadding="2" cellspacing="1" width="95%" align="center" bgcolor="#6699ff">
  <tr align="center"> 
    <td bgcolor="#FFFFFF" valign="bottom" height="18" colspan="6"><b><%=rs("iname")%>��ϸ����</b><b>>>></b> 
    </td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" valign="bottom" height="18" width="171"><font  color="#000099">����:</font></td>
    <td bgcolor="#FFFFFF" valign="bottom" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" valign="bottom" height="18" width="215"><%=rs("iname")%></td>
    <td bgcolor="#FFFFFF" valign="bottom" height="18" width="121" align="right"><font  color="#000099">�Ա�:</font></td>
    <td bgcolor="#FFFFFF" valign="bottom" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" valign="bottom" height="18" width="238"><%=rs("sex")%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="18" width="171"> 
      <p align="right"><font  color="#000099">��������:</font></p>
    </td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("bday")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">���֤����:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"></td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("code")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">����:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("mzhu")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">����״��:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("marry")%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="18" width="171"> 
      <p align="right"><font  color="#000099">�������ڵ�:</font></p>
    </td>
    <td bgcolor="#FFFFFF" height="18" width="22">��</td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("huji")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">��ס��:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("szd")%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="18" width="171"> 
      <p align="right"><font  color="#000099">��ߵĽ����̶�:</font></p>
    </td>
    <td bgcolor="#FFFFFF" height="18" width="22">��</td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("edu")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">רҵ:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("zye")%></td>
  </tr>
  <tr> 
    <td height="18" bgcolor="#FFFFFF" width="171"> 
      <p align="right"><font  color="#000099">��ҵԺУ:</font></p>
    </td>
    <td height="18" bgcolor="#FFFFFF" width="22">��</td>
    <td height="18" bgcolor="#FFFFFF" width="215"><%=rs("school")%></td>
    <td height="18" bgcolor="#FFFFFF" width="121" align="right"><font  color="#000099">ְҵ:</font></td>
    <td height="18" bgcolor="#FFFFFF" width="23"> </td>
    <td height="18" bgcolor="#FFFFFF" width="238"><%=rs("work")%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="18" width="171"> 
      <p align="right"><font  color="#000099">������ò:</font></p>
    </td>
    <td bgcolor="#FFFFFF" height="18" width="22">��</td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("zzmm")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font s color="#000099">����ְ��:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("zchen")%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="18" width="171"><font color="#000099"></font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font color="#000099"></font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"></td>
    <td bgcolor="#FFFFFF" height="18" width="238"> </td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">��ϵ�绰:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("phone")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font " color="#000099">�ֻ�/����:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("callnum")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">�����ʼ�:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("email")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font color="#000099">��ϵ��ַ:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("address")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" colspan="6"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td width="100%" height="5">..............................................................................................................</td>
        </tr>
        <tr> 
          <td width="100%" height="10"></td>
        </tr>
        <tr> 
          <td width="100%" height="20"><b>��ְ��Ϣ>>></b></td>
        </tr>
      </table>
    </td>
  </tr>
  <%
uname_=rs("uname")
iname_=rs("iname")
 rs.close
set rs=nothing
 Set rs = Server.CreateObject("ADODB.Recordset")
sql2="select * from qiuzhi where uname='"&uname_&"' and iname='"&iname_&"'"
rs.open sql2,conn,1,1
if rs.eof then
response.write "<tr><td>������ְ��Ϣ��</td></tr></table>"
response.end
end if
%>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font color="#000099">Ѱ��ְλ:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("biaoti")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"> </td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"> </td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">�����س�:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("waiyu")%> �ȼ���<%=rs("dengji")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">���������</font>:</td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("computer")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font color="#000099">��Ҫ����:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" colspan="4"><%=rs("jineng")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">���ܾ���:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215">������ؾ��鹲��<%=rs("gznum")%> ��</td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"> </td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"> </td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">��������:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" colspan="4"><%=rs("gzjl")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">����ְ�������ҵ��ģ:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("qyrs")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">����ְ��:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("zhiwu")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">��ʤ�ι���:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" colspan="4"><%=rs("job")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">ϣ�������ص�:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" width="215"><%=rs("gzdd")%></td>
    <td bgcolor="#FFFFFF" height="18" width="121" align="right"><font  color="#000099">н��Ҫ��:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="23"> </td>
    <td bgcolor="#FFFFFF" height="18" width="238"><%=rs("yuex")%></td>
  </tr>
  <tr> 
    <td align="right" bgcolor="#FFFFFF" height="18" width="171"><font  color="#000099">����Ҫ��:</font></td>
    <td bgcolor="#FFFFFF" height="18" width="22"></td>
    <td bgcolor="#FFFFFF" height="18" colspan="4"><%=rs("qtyq")%></td>
  </tr>
  <tr> 
    <td width="171" bgcolor="#FFFFFF"></td>
    <td width="22" bgcolor="#FFFFFF"></td>
    <td width="215" bgcolor="#FFFFFF"></td>
    <td width="121" bgcolor="#FFFFFF"></td>
    <td width="23" bgcolor="#FFFFFF"></td>
    <td width="238" bgcolor="#FFFFFF"></td>
  </tr>
</table>
                                          </body>
</html>
<%
rs.close
set rs=nothing
%>