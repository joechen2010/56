<%@ codepage ="936" %>
<!--#include file="../../Inc/Conn.asp" -->
<LINK href="../style.css" rel=stylesheet type=text/css> 
<%
dim sql
dim rs
sql="select * from qyml where id="&request("id")
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 

      
<TABLE border=0 cellspacing=1 width="96%" style="border-collapse: collapse" cellpadding="6" align="center" bgcolor="#e8f4ff">
  <TBODY> 
  <tr bgcolor="#e8f4ff"> 
    <TD colspan=2 height=30 width="460"><b>�޸�ע���Ա��Ϣ</b></TD>
  </tr>
  <TR bgcolor="#FFFFFF"> 
    <TD align="center">����Ա��¼����</TD>
    <TD width=312><%=rs("user")%></TD>
  </TR>
  <TR bgcolor="#FFFFFF"> 
    <TD align="center">���� �룺</TD>
    <TD width="312"><%=rs("pass")%> </TD>
  </TR>
  <TR bgcolor="#FFFFFF"> 
    <TD align="center">��������ʾ����</TD>
    <TD width="312"><%=rs("question")%></TD>
  </TR>
  <TR vAlign="middle" bgcolor="#FFFFFF"> 
    <TD align="center">��������ʾ�𰸣�</TD>
    <TD width="312"><%=rs("answer")%></TD>
  </TR>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">������������</TD>
    <TD width="482"><%=rs("name")%> <%=rs("ch")%> </TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center" height="19">������ְ��</TD>
    <TD width="312"><%=rs("zw")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">����&nbsp;&nbsp;&nbsp; �ţ�</TD>
    <TD width="312"><%=rs("bm")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">���ֵ���ַ��</TD>
    <TD width="312"><%=rs("address")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">����&nbsp;&nbsp;&nbsp; �ࣺ</TD>
    <TD width="312"><%=rs("post")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">��ʡ��������</TD>
    <TD width="312"><%=rs("province")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">����&nbsp;&nbsp;&nbsp; �У�</TD>
    <TD width="312"><%=rs("city")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">���绰��</TD>
    <TD width="312"> 
      <%response.write rs("phone1")&"-"&rs("phone2")&"-"&rs("phone3")%>
    </TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">����&nbsp;&nbsp;&nbsp; �棺</TD>
    <TD width="312"> 
      <%response.write rs("fax1")&"-"&rs("fax2")&"-"&rs("fax3")%>
    </TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">�����ĵ����ʼ���</TD>
    <TD width="312"><%=rs("email")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD align="center">����&nbsp;&nbsp;&nbsp; ַ��</TD>
    <TD width="312"><%=rs("web")%> </td>
  </tr>
  <tr bgcolor="#e8f4ff"> 
    <TD colspan=2 height=30 width="460"><b>�޸Ĺ�˾��Ϣ</b></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">��˾���ƣ�</TD>
    <TD width="312"><%=rs("qymc")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">���˴���</TD>
    <TD width="312"><%=rs("frdb")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">��˾���ʣ�</TD>
    <TD width="312" ><%=rs("qyxz")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">ע���ʽ�</TD>
    <TD width="312"><%=rs("zczj")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">��Ҫ��Ʒ��</TD>
    <TD width="312"><%=rs("zycp")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">��Ա������</TD>
    <TD width="312"><%=rs("ygrs")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">��Ӫҵ�</TD>
    <TD width="312"><%=rs("nyye")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD width="145" align="center">��˾���ܣ�</TD>
    <TD width="312"><%=rs("qyjj")%></TD>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <TD height="10" width="145" ></TD>
    <TD height="10" width="145" ></TD>
  </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <TD height="20" colspan="2"><a href="javascript:window.close()">�ر�</a></TD>
  </tr>
  </TBODY> 
</TABLE>
