<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="check.asp"-->
<%
dim subject,content,F_Company,F_Name,F_Email,F_address,F_Post,F_Tel,F_Fax,T_User,F_User,TF_Time
subject=request("subject")
content=request("content")
F_Company=request("F_Company")
F_Name=request("F_Name")
F_Email=request("F_Email")
F_address=request("F_address")
F_Post=request("F_Post")
F_Tel=request("F_Tel")
F_Fax=request("F_Fax")
T_User=request("T_User")
F_User=request("F_User")
T_Name=request("T_Name")
T_Company=request("T_Company")
T_Email=request("T_Email")
set rs=server.createobject("adodb.recordset")
sql="select * from message"
rs.open sql,conn,1,3
rs.addnew
rs("subject")=subject
rs("content")=content
rs("F_Company")=F_Company
rs("F_Name")=F_Name
rs("F_Email")=F_Email
rs("F_address")=F_address
rs("F_Post")=F_Post
rs("F_Tel")=F_Tel
rs("F_Fax")=F_Fax
rs("T_User")=T_User
rs("F_User")=F_User
rs("T_Name")=T_Name
rs("T_Email")=T_Email
rs("T_Company")=T_Company
rs("TF_Time")=now()
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<TITLE>���������� ��Ա�������� - �������˵����ϼ�԰</TITLE>
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
<body>
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
            <a href="http://www.cx56w.com">����������</a> &gt; �޸Ĺ�˾����</td>
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
                        <div align="center"><font color="#FFFFFF">�޸Ĺ�˾����</font></div>
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
                  <p align=center>��Ϣ���ͳɹ���</p>
                  <p> �����ǳ����������Ļ�Ա�����ھͼ��룬���ܸ����Ա���� * �����Ʒ��ʲ�Ʒ����Ӧ�̵����ݿ� * �ҵİ칫�����ɹ�������ó�ס��շ���Ϣ���޸����� 
                    * ֱ����ϵ�̼ң�������ҵ��Ϣ�����Ĳ�Ʒ�ٵ� * �˽����ó�׷��ɷ��� * ��ú�����ó�׷�չ״�� �����ǳ����������Ļ�Ա������ע�ᣬ�������ڳ����������Ͻ�Ҫ�շ���������Ϣ�� 
                    �Ѿ��ǳ����������Ļ�Ա�����˵�¼���������ڳ����������Ͻ�Ҫ�շ���������Ϣ�� </p>
                  <p align=center><a href="javascript:window.history.back()">����</a></p>
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
