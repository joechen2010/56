<%@ Language=VBScript codepage ="936" %>
<!--#include file="../inc/conn.ASP"--> 
<HTML> 
<head>
<TITLE>ȡ������</TITLE>
<LINK href="images/1css.css" type="text/css" rel="stylesheet">
<style type="text/css">.style1 { FONT-WEIGHT: bold; COLOR: #ff8200 }
	.style2 { COLOR: #ff8200 }
	.p14 { FONT-SIZE: 14px; COLOR: #000000; LINE-HEIGHT: 130% }
	</style>
</head>
<body leftmargin="0" topmargin="0">
<!--#include file="Head.htm"-->
<table cellspacing="0" cellpadding="0" width="770" align="center" bgcolor="#ff8200" border="0">
  <tr> 
    <td height="2"></td>
  </tr>
  <tr> 
    <td bgcolor="#c66300" height="1"></td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="770" align="center" border="0">
  <tr> 
    <td class="S" height="30">����ǰλ�ã�<a href="../">��ҳ</a> &gt;&gt;&nbsp; ��Ա��¼ </td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="98%" align="center" border="0">
  <tr> 
    <td height="30">&nbsp;</td>
  </tr>
</table>
<TABLE cellSpacing=0 cellPadding=0 width=760 align=center border=0 class=S>
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
                              color=#087dbd>��ӭ��������������!</FONT></STRONG></DIV></TD></TR></TBODY></TABLE>
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
                                <TD class=S>&nbsp;&nbsp;&nbsp;&nbsp;�������ʹ������ʲô���ʻ�����ǵķ�����ʲô�ܺõĽ��飬����ʱ��ϵ���ǡ� 
                                <BR>
                                              �绰�� 0574-63809664<BR>
                                              ���棺 0574-63897836<BR>
                                              ���䣺 <A 
                                href="mailto:service@gbearing.com">service@gbearing.com</A></TD>
                                          </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></DIV></TD>
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
                              color=#087dbd>�������룿��ȷ��һ�����!</FONT></FONT></TD></TR>
                          <TR>
                            <TD class=LL align=middle><BR>
                              <DIV id=fri><SPAN id=saynow></SPAN><BR>
                              <table border="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1" height="144"> 
<tr> 
                                      <td height="128" align="center"> 
                                        <%if request("action")="" then%>
                                        <FORM ACTION="getAnswer.ASP" method="post" name=getanswer> 
                                          <p align="center">�û����� 
                                            <input type=text name=UserNameGet size="20">
                                            <input type=hidden name=Sure value="sure">
                                            <input type=hidden name=action value="action">
                                            <input type=submit name=submit value="ȡ������" style="border-style: solid; border-width: 1">
                                          </p>
                                        </form>
<%elseif request("action")="action" then%>
<FORM ACTION="getPassword.ASP" method="post" name=getanswer>
<%
dim UsernameGet
UsernameGet=Request.Form("UsernameGet")
sql="select * from qyml where User='"&UsernameGet&"'"
set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1
if RS.eof and rs.bof then
   Response.Write "���û������ڣ�"
'   response.end
else
Response.Write Request.Form("UsernameGet")
Response.Write "�������ǣ�"
Response.Write RS("Question") 
%><br>
<input type="hidden" name=UserName value=<%=Request.Form("UserNameGet")%>>
                                          �� ���� 
                                          <input type=text name=Answer size="20">
                                          &nbsp;
                                          <input type=submit name=submit value="�һ�����" style="border-style: solid; border-width: 1">
                                        </form>
<%
end if
else
    response.write "�������"
end if%></td></tr> <tr> <td width="100%" height="12"></td></tr> </table><BR>
                              <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                              border=0>
                                <TBODY>
                                <TR>
                                <TD height=25>
                                        <DIV class=HEADWHITE2 align=center><STRONG>������<A 
                                href="../reg/User_Reg.asp" 
                                target=_blank><FONT 
                                color=#ff6600>�������</FONT></A>���½���ע��</STRONG></DIV>
                                      </TD></TR></TBODY></TABLE></DIV>
                                <SPAN 
                              class=font12><STRONG><FONT 
                              color=#ff6600>�����Դ�绰����0574-63809664--�ͷ����� Ҫ���ʹ�����롣</FONT></STRONG></SPAN> 
                                <P align=center>
                              <P></P></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
                  </TD>
                </TR></TBODY></TABLE><BR></DIV></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<!--#include file="../Inc/down.htm"-->
</BODY> 
</HTML>