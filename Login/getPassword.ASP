<%@ Language=VBScript codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<HTML> 
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
<LINK href="images/1css.css" type="text/css" rel="stylesheet">
<style type="text/css">.style1 { FONT-WEIGHT: bold; COLOR: #ff8200 }
	.style2 { COLOR: #ff8200 }
	.p14 { FONT-SIZE: 14px; COLOR: #000000; LINE-HEIGHT: 130% }
	</style>
<TITLE>取回密码</TITLE> 
</HEAD>
<%
dim UsernameGet,PasswordGet
dim myConn,myRS
UsernameGet=Request.Form("Username")
AnswerGet=Request.Form("Answer")
sql="select * from qyml where User='"&UsernameGet&"' and answer='"&answerget&"' "
set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1
%>
<BODY leftmargin="0" topmargin="0">
<!--#include file="Head.htm"-->
<table cellspacing="0" cellpadding="0" width="770" align="center" bgcolor="#ff8200" border="0">
  <tr> 
    <td height="2"></td>
  </tr>
  <tr> 
    <td bgcolor="#c66300" height="1"></td>
  </tr>
</table>
<TABLE cellSpacing=0 cellPadding=0 width=760 align=center border=0>
  <TBODY>
  <TR>
    <TD>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
        <TBODY>
        <TR>
          <TD vAlign=top>
            <DIV align=center>
            <TABLE height=300 cellSpacing=1 cellPadding=0 width=760 align=center 
            bgColor=#cccccc border=0>
              <TBODY>
              <TR>
                  <TD vAlign=top width=55 bgColor=#ffffff height=110><IMG 
                  height=260 
                  src="images/register_left_1.gif" 
                  width=55 border=0></TD>
                  <TD vAlign=top 
                background=images/register_top_1.gif 
                bgColor=#ffffff> 
                    <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                        <TD colSpan=3><IMG height=80 
                        src="images/login_banner.gif" 
                        width=702></TD>
                      </TR>
                    <TR>
                      <TD vAlign=top width="38%">
                        <DIV align=center>
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY>
                          <TR>
                            <TD height=30>
                              <DIV align=center><STRONG><FONT 
                              color=#087dbd>欢迎来到诚信物流网!</FONT></STRONG></DIV></TD></TR></TBODY></TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY>
                          <TR>
                            <TD>&nbsp;</TD></TR>
                          <TR>
                            <TD>
                                  <TABLE height=118 cellSpacing=0 cellPadding=0 
                              width=243 align=center 
                              background=images/login_bg.gif 
                              border=0>
                                    <TBODY> 
                                    <TR>
                                <TD>
                                <TABLE cellSpacing=0 cellPadding=0 width="95%" 
                                align=center border=0>
                                <TBODY>
                                <TR>
                                <TD><SPAN 
                                class=font12>&nbsp;&nbsp;&nbsp;&nbsp;如果您在使用中有什么疑问或对我们的服务有什么很好的建议，请随时联系我们。 
                                <BR>
                                              电话： 0574-63809664<BR>
                                              传真： 0574-63897836<BR>
                                              邮箱： <A 
                                href="mailto:service@gbearing.com">service@gbearing.com</A></SPAN></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></DIV></TD>
                        <TD vAlign=center width="1%" 
                      background=images/xian.gif> 
                          <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY>
                          <TR>
                            <TD>&nbsp;</TD></TR></TBODY></TABLE></TD>
                      <TD vAlign=top width="61%">
                        <TABLE cellSpacing=0 cellPadding=6 width="100%" 
border=0>
                          <TBODY>
                          <TR>
                            <TD class=HEADWHITE2 align=middle 
                              bgColor=#f1f1f1><FONT color=#ff6600><FONT 
                              color=#087dbd>忘记密码？先确认一下身份!</FONT></FONT></TD></TR>
                          <TR>
                            <TD class=LL align=middle><BR>
                              <DIV id=fri><SPAN id=saynow></SPAN><BR>
                              <BR><table border="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1" height="144"> 
<tr> 
                                      <td width="100%" height="128" align="center"> 
                                        <%
if rs.eof and rs.bof then
Response.Write "<p>您输入的答案错误！<a href='javascript:history.go(-1)'>请重新输入</a><BR>"
else
Response.Write "您的会员号是："&RS("User")&"<br><br>"
Response.Write "您的密码："&RS("pass")
end if
%>
                                        &nbsp; </td>
                                    </tr> </table>
                              <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                              border=0>
                                <TBODY>
                                <TR>
                                <TD height=25>
                                        <DIV class=HEADWHITE2 align=center><STRONG>您可以<A 
                                href="../reg/User_Reg.asp" 
                                target=_blank><FONT 
                                color=#ff6600>点击这里</FONT></A>重新进行注册</STRONG></DIV>
                                      </TD></TR></TBODY></TABLE></DIV>
                                <SPAN 
                              class=font12><STRONG><FONT 
                              color=#ff6600>您可以打电话到：0574-63809664--客服热线 要求初使化密码。</FONT></STRONG></SPAN> 
                                <P align=center>
                              <P></P></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
                  </TD>
                </TR></TBODY></TABLE><BR></DIV></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<!--#include file="../Inc/down.htm"-->
</BODY> 
</HTML>