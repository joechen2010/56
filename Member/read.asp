<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<%
conn.execute("update message set new=1 where id="&request("id"))
set rs=server.createobject("adodb.recordset")
sql="select * from message where id="&request("id")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
    response.write "û�����ݣ���<a href=""javascript:windows.history.back()"">����</a>"
else
%>
<html>
<head>
<TITLE>���������� ��Ա�������� - �������˵����ϼ�԰</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {
	background-color: #2C68B1;
	margin: 0px;
}
-->
</style>
<link href="images/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<!--#include file="head.htm"-->
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">����������</a> &gt; �յ�������Ϣ</td>
          <td width="17%"><img src="images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="534"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> <%=session("user")%> </font> </strong>����ӭ��������</td>
        </tr>
      </table>
      <table width="100%" height="497"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top"><table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="images/title_2.jpg" bgcolor="#FFFFFF" >
                        <div align="center"><font color="#FFFFFF"> �յ�������Ϣ</font></div>
                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
			  <table width="100%"  border="0">
                             <tr>
                              <td align="left"><div id="TipAll" >
                               ���������б���*����Ŀ��
                              </div></td>
                            </tr>
                           </table>
                  <table cellspacing=1 cellpadding=3 width=100% align=center 
              border=0 bgcolor="#B1D4F2">
                    <tbody> 
                    <tr bordercolor=#ffffcc bgcolor="#EDF6FF"> 
                      <td width=100><b>����:</b></td>
                      <td><b><%=rs("subject")%></b></td>
  </tr>
                    <tr bgcolor="#EDF6FF"> 
                      <td width=100 height=28>�����˹�˾��</td>
                      <td><%=rs("F_Company")%></td>
  </tr>
                    <tr bgcolor="#EDF6FF"> 
                      <td width=100 height=28>�����ˣ�</td>
                      <td><%=rs("F_name")%></td>
  </tr>
                    <tr valign=top bgcolor="#EDF6FF"> 
                      <td valign=top width=100 
                  height=120>���ݣ�<br>
    </td>
                      <td height=120><%=dvHTMLEncode(rs("content"))%></td>
  </tr>
                    <tr valign=top bgcolor="#EDF6FF"> 
                      <td valign=top width=100 
                height=28>�����˵�ַ��</td>
                      <td><%=rs("F_Address")%>���ʱࣺ<%=rs("F_Post")%>��</td>
  </tr>
  <!--<tr valign="top">
<td width="100" valign=top bgcolor="#DDE8F0" height="9">Mail:</td>
<td height="-2" bgcolor="#F0F8F9">jsbamboo@jsbamboo.net</td>
</tr>-->
                    <tr valign=top bgcolor="#EDF6FF"> 
                      <td valign=top width=100 
                height=28>�绰��</td>
                      <td><%=rs("F_Tel")%></td>
  </tr>
                    <tr valign=top bgcolor="#EDF6FF"> 
                      <td valign=top width=100 height=28>���棺</td>
                      <td><%=rs("F_Fax")%></td>
  </tr>
                    <tr valign=top bgcolor="#EDF6FF"> 
                      <td width=100 height=28>Email��</td>
                      <td><%=rs("F_Email")%></td>
  </tr>
                    <tr bgcolor="#EDF6FF"> 
                      <td></td>
                      <td align=left> 
                        <table cellspacing=0 cellpadding=4 width=300>
                          <tbody> 
                          <tr align="center"> 
          <td width=100><a href="reply.asp?id=<%=rs("id")%>">�ظ�</a></td>
          <td width=100><a href="javascript:window.history.back()">����</a></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  </tbody> 
</table> 
				</td>
            </tr>
          </table>
                        <table width="534"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="right"><a href="#top"><img src="images/icon_top.gif" alt="�ص�ҳ�涥��" border="0" align="absmiddle" width="15" height="18"></a> 
                  <a href="#top">�ټ��һ��</a></td>
              </tr>
            </table></td>
        </tr>
      </table>
</td>
  </tr>
</table>

<!--#include file="bottom.htm"-->
</body>
</html>
<%end if%>
