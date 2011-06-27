<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../Member/check.asp"-->

<html>
<head>
<TITLE>诚信物流网 会员控制中心 - 物流商人的网上家园</TITLE>
<style type="text/css">
<!--
body {
	background-color: #2C68B1;
	margin: 0px;
}
-->
</style>
<link href="../Member/images/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="566" align="center" valign="top" bgcolor="#FFFFFF" id="main">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="../Member/images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 回复留言</td>
          <td width="17%"><img src="../Member/images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="100%" height="497"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top"><table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="../Member/images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="../Member/images/title_2.jpg" bgcolor="#FFFFFF" >
                        <div align="center"><font color="#FFFFFF">回复留言</font></div>                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="../Member/images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
			  <table width="100%"  border="0">
<%set rs2=server.createobject("adodb.recordset")
sql2="select * from qyml where user='"&request("T_user")&"'"		
rs2.open sql2,conn,1,1	
if rs2.eof and rs2.bof then
    response.write "没有数据"
else		  
%>                       <td colspan="2" align="center">我给<font color="#FF0000"><%=rs2("qymc")%></font>的留言</td>
<%end if%>
                  </table>
                  <TABLE border=0 cellPadding=5 cellSpacing=1 width="100%" align="center" bgcolor="#B1D4F2">
                    <form method="POST" action="tj_save.asp?T_User=<%=request("T_User")%>" name="Form1">
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"><font color="#FF0000">* </font><font color="#000000"></font><font color="#000000">主题：</font></TD>
                        <TD width="503"> 
                          <input type="text" name="subject"></TD>
                      </TR>
                      <TR bgcolor="#EDF6FF"> 
                        <TD width="96" align="right"><font color="#FF0000">* </font>内容：</TD>
                        <TD width="503" bgcolor="#EDF6FF"> 
                        <textarea cols=50 name=content rows=6 style="border-style: solid; border-width: 1"></textarea>                        </TD>
                      </TR>
                      <TBODY>
                      <TR ALIGN="CENTER" bgcolor="#EDF6FF"> 
                        <TD COLSPAN="2"> 
                          <INPUT TYPE="submit" VALUE="确 定" NAME="button">&nbsp;
                          <INPUT NAME=reset CLASS="smallinput" TYPE="button" VALUE="返 回" ONCLICK=javascript:window.close()>                        </TD>
                      </TR>
                      </TBODY>
                    </form>
                  </TABLE>                </td>
            </tr>
          </table>
                        </td>
        </tr>
      </table></td>
  </tr>
</table>

</body>
</html>

