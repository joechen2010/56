<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
dim url
dim errmsg
url=request.form("url")
set rs=server.CreateObject ("adodb.recordset")
sql="select * from qyml where id="&gsid
rs.open sql,conn,1,3
rs("url")=url
rs.update
session("url")=rs("url")
rs.close
set rs=nothing 
conn.close
set conn=nothing 
%>

<HTML>
<HEAD>
<TITLE>���������� ��Ա�������� - �������˵����ϼ�԰</TITLE>
<META content=zh-cn http-equiv=Content-Language> 
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
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
</HEAD>
<BODY>
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
            <a href="http://www.cx56w.com">����������</a> &gt; ��վ�Զ�����</td>
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
                        <div align="center"><font color="#FFFFFF">��վ�Զ�����</font></div>
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
                  <table border="0" cellspacing="1" width="100%" cellpadding="5" style="border-collapse: collapse" bgcolor="#FF9900">
                    <tr> <td height="431" bgcolor="#FFFBEE" width="100%"><P style="line-height: 180%"><b> 
<span style="font-size: 10.5pt">&nbsp;&nbsp; </span></b> <font color="#FF6600"> 
<span style="font-size: 10.5pt"><b><br> &nbsp;��ϲ��������վ�Ѿ��Զ����ɣ�</b></span></font></P><div align="center"> 
<center> <table border="0" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="99%" id="AutoNumber11" height="57"> 
<tr> <td width="18%" height="26">�����ַ��:</td>
                                <td width="82%" height="26"> <a target="_blank" href="<%=WebUrl%>/showroom/index.asp?gsid=<%=gsid%>"> 
                                  <%=WebUrl%>/showroom/index.asp?gsid=<%=gsid%></a></td>
                              </tr> <tr> <td width="18%" height="11">���Ӵ���:</td><td width="82%" height="11"> 
<input type=text size=54 value="&lt;a href=<%=WebUrl%>/showroom/index.asp?gsid=<%=gsid%>&gt;<%=companyname%>&lt;/a&gt;" name="a"></td></tr> 
<tr> <td width="100%" height="5" colspan="2">�����԰Ѹ���ַӡ����Ƭ�ϡ���ҵ�����϶��������ˡ�</td></tr> 
</table></center></div><table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" height="211"> 
<tr> <td align="center" height="116"><font color="#008000"><br> <br> </font> <b><font style="font-size: 10.5pt"> 
<a target="_blank" href="../../showroom/index.asp?gsid=<%=gsid%>">��˽���&gt;&gt;�ҵ���վ</a></font></b><font color="#008000"><span style="font-size: 10.5pt"><br> 
</span> <br> <br> <br> </font> ��ӭ����-=&gt;����Ϊ���ṩ�������ʶ�������<br> <font color="#008000"><br> 
</font> </td></tr> <tr> <td height="12"><font color="#FF9900">��</font>����������������������:</td></tr> 
<tr> 
          <td height="12">&nbsp; [1]:����������:http://www.cx56w.com : http://www.gbearing.com 
            ,�ȹ��ʹ��ڶ�������</td>
        </tr> 
<tr> <td height="12">���ǰ������������շѱ�׼�����շѡ�</td></tr> <tr> 
          <td height="12">&nbsp; [2]:����Ϊ��������:http://www.bearingpage.net/company 
            : http://yourname.bearingpage.net/</td>
        </tr> <tr> <td height="23">��������ṩ</td></tr> </table></td></tr> 
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
