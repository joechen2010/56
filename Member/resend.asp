<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from message where id="&request("id")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
    response.write "û�����ݣ���<a href=""javascript:windows.history.back()"">����</a>"
	response.end
else
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
            <a href="http://www.cx56w.com">����������</a> &gt; �ط��ظ���Ϣ</td>
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
                        <div align="center"><font color="#FFFFFF">�ط��ظ���Ϣ</font></div>
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
                  <TABLE border=0 cellPadding=5 cellSpacing=1 width="100%" align="center" bgcolor="#B1D4F2">
                    <form method="POST" action="messagesave.asp" name="Form1">
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right">�ռ��˹�˾��</TD>
                                        
                        <TD width="503"><b class="L"><%=rs("T_Company")%></b></TD>
                                      </TR>
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"><font color="#000000">�ռ��ˣ�</font></TD>
                                        
                        <TD width="503"><%=rs("T_name")%></TD>
                                      </TR>
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"><font color="#FF0000">* </font><font color="#000000"></font><font color="#000000">���⣺</font></TD>
                                        
                        <TD width="503"> 
                          <input type="text" name="subject" value="<%=rs("subject")%>">
                                        </TD>
                                      </TR>
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"><font color="#FF0000">* </font>���ݣ�</TD>
                                        
                        <TD width="503"> 
                          <textarea cols=50 name=content rows=6 style="border-style: solid; border-width: 1"><%=rs("content")%></textarea>
                                        </TD>
                                      </TR>
                                      <TBODY> 
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"><FONT color=#ff6600>*</FONT> 
                          ���Ĺ�˾��</TD>
                                        
                        <TD width="503"> 
                          <input type="text" size="33" name="F_Company" value=<%=rs("F_Company")%>>
                                        </TD>
                                      </TR>
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"><FONT color=#ff6600>*</FONT> 
                          �������֣�</TD>
                                        
                        <TD width="503"> 
                          <input type="text" size="33" name="F_Name" value=<%=rs("F_Name")%>>
                                        </TD>
                                      </TR>
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"><font color=#ff6600>*</font> 
                          �����ʼ���</TD>
                                        
                        <TD width="503"> 
                          <input type="text" size="33" name="F_Email" value=<%=rs("F_Email")%>>
                                        </TD>
                                      </TR>
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"><FONT color=#ff6600>*</FONT> 
                          ��ϵ��ַ��</TD>
                                        
                        <TD width="503"> 
                          <input type="text" size="33" name="F_address" value=<%=rs("F_address")%>>
                                        </TD>
                                      </TR>
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"><FONT color=#ff6600>*</FONT> 
                          ��&nbsp;&nbsp;&nbsp;&nbsp;��:</TD>
                                        
                        <TD width="503"> 
                          <input type="text" size="33" name="F_Post" value=<%=rs("F_Post")%>>
                                        </TD>
                                      </TR>
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"> <font color=#ff6600>*</font> 
                          ��&nbsp;&nbsp;&nbsp;&nbsp;����</TD>
                                        
                        <TD width="503"> 
                          <input type="text" size="33" name="F_Tel" value=<%=rs("F_Tel")%>>
                                        </TD>
                                      </TR>
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right">&nbsp; <font color=#ff6600>*</font> 
                          ��&nbsp;&nbsp;&nbsp;&nbsp;��:</TD>
                                        
                        <TD width="503"> 
                          <input type="text" size="33" name="F_Fax" value=<%=rs("F_Fax")%>>
                                        </TD>
                                      </TR>
                                      
                      <TR bgcolor="#EDF6FF"> 
                        <TD COLSPAN="2"> 
                          <input type="hidden" name="F_User" value="<%=rs("F_User")%>">
                                          <input type="hidden" name="T_User" value="<%=rs("T_User")%>">
										  <input type="hidden" name="T_Name" value="<%=rs("T_Name")%>">
                                          <input type="hidden" name="T_Company" value="<%=rs("T_Company")%>">
										  <input type="hidden" name="T_Email" value="<%=rs("T_Email")%>">
                                        </TD>
                                      </TR>
                                      
                      <TR ALIGN="CENTER" bgcolor="#EDF6FF"> 
                        <TD COLSPAN="2"> 
                          <INPUT TYPE="submit" VALUE="ȷ ��" NAME="button">
                                          <IMG HEIGHT=1 WIDTH=80> 
                                          <INPUT NAME=reset CLASS="smallinput" TYPE="button" VALUE="�� ��" ONCLICK=javascript:history.go(-1)>
                                        </TD>
                                      </TR>
                                      </TBODY></form> 
                                    </TABLE>
                                  
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
