<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../Inc/Function.asp"-->
<!--#include file="../inc/config.asp"-->
<HTML>
<HEAD>
<TITLE><%=WebName%>[��Ա��¼]</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
<SCRIPT language=JavaScript>
<!--
function docheck(){
if (loginForm.UsernameGet.value==""){
alert("�������û���!");
loginForm.UsernameGet.focus();
return false;
}
if (loginForm.PasswordGet.value==""){
alert("����������!");
loginForm.PasswordGet.focus();
return false;
}
loginForm.UsernameGet.value=loginForm.UsernameGet.value.toLowerCase();
//loginForm.submit();
return true;
}
//-->
</SCRIPT>

<style type="text/css">
<!--
body {
	background-color: #2C68B1;
	margin: 0px;
}
-->
</style>
<link href="../member/images/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
</HEAD>
<BODY text=#000000 leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="../inc/top.htm"-->
<table cellspacing="0" cellpadding="0" width="770" align="center" bgcolor="#ff8200" border="0">
  <tr> 
    <td height="2"></td>
  </tr>
  <tr> 
    <td bgcolor="#c66300" height="1"></td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="778" align="center" border="0" bgcolor="#FFFFFF">
  <tr> 
    <td class="S" height="2"></td>
  </tr>
</table>
<TABLE width=778 border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <TBODY> 
  <TR> 
    <TD vAlign=top width=160 bgColor=#2c68b1 height=566> 
      <!--#include file="left.htm"-->
    </TD>
    <TD id=main align=middle bgColor=#ffffff valign="top"> 
      <table cellspacing=0 cellpadding=3 width="94%" border=0>
        <tbody> 
        <tr> 
          <td id=pos align=left width="83%" height=38><img height=11 
            src="images/icon03.gif" width=15> <a 
            href="../">��ҳ</a> &gt; ��Ա����������ҳ</td>
          <td width="17%"><img height=32 
            src="images/icon_onlineservice.gif" 
      width=132></td>
        </tr>
        </tbody> 
      </table>
      <br>
      <table cellspacing=0 cellpadding=0 width="90%" border=0>
        <tbody> 
        <tr> 
          <td width=7 height=6><img height=6 src="images/con_1.gif" 
          width=7></td>
          <td background=images/con_bg1.gif></td>
          <td width=6 height=6><img height=6 src="images/con_2.gif" 
          width=6></td>
        </tr>
        <tr> 
          <td background=images/con_bg2.gif></td>
          <td id=sysmsg valign=top align=left 
            bgcolor=#feffd2>����δ��¼�����������������������û������룬�������¼����<br>
            �����û��ע�������������Ա���밴�������ʾ���<a 
            href="../reg/User_Reg.asp">ע��</a>��</td>
          <td background=images/con_bg4.gif></td>
        </tr>
        <tr> 
          <td width=7 height=6><img height=6 src="images/con_3.gif" 
          width=7></td>
          <td background=images/con_bg3.gif></td>
          <td width=6 height=6><img height=6 src="images/con_4.gif" 
          width=6></td>
        </tr>
        </tbody> 
      </table>
      <br>
      <div id=sysmsg> 
        <div id=login></div>
      </div>
      <form action=login_check.asp method=post>
        <table cellspacing=0 cellpadding=5 width=330 align=center bgcolor=#8bc1cd 
      border=0>
          <tbody> 
          <tr> 
            <td class=tbr11 align=middle width=122 bgcolor=#b9d5f7><font 
            class=fs14bb>��ע���Ա��¼</font></td>
            <td class=tbr21 align=middle width=188 bgcolor=#edf6ff>&nbsp;<a 
                  href="getAnswer.asp">�������룿</a></td>
          </tr>
          <tr align=middle bgcolor=#b9d5f7> 
            <td class=tbr31 colspan=2> 
              <table cellspacing=1 cellpadding=3 width="100%" align=center 
border=0>
                <tr bgcolor="#edf6ff"> 
                  <td colspan="2"><strong><font 
                  color=#990000>���ǻ�Ա��</font></strong></td>
                </tr>
                <tbody> 
                <tr bgcolor="#edf6ff"> 
                  <td><img height=30 src="images/1ic_gal_01.gif" 
                width=29></td>
                  <td><font 
color=#ff0000>�����������û��������룡</font></td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=1 cellpadding=3 align=center border=0 width="100%">
                <tbody> 
                <tr bgcolor="#edf6ff"> 
                  <td align=right height=25> 
                    <div align=center><strong>��Ա����</strong> </div>
                  </td>
                  <td height=25> 
                    <input id=TextBox1 style="WIDTH: 120px" 
                  name=UserID>
                  </td>
                </tr>
                <tr bgcolor="#edf6ff"> 
                  <td align=right height=25> 
                    <div align=center><strong>��&nbsp;&nbsp;�룺</strong> </div>
                  </td>
                  <td height=25> 
                    <input id=TextBox2 style="WIDTH: 120px" 
                  type=password name=Pwd>
                  </td>
                </tr>
                <tr bgcolor="#edf6ff"> 
                  <td align=right colspan=2 height=25> 
                    <div align=center> 
                      <input type="hidden" name="action" value="<%=request("action")%>">
                      <%url=Request.ServerVariables("HTTP_REFERER")%>
                      <input type="hidden" name="url" value="<%=url%>">
                      <input id=Image1 type=image height=32 width=78 src="images/1login.gif" border=0 name=Image1>
                    </div>
                  </td>
                </tr>
                </tbody> 
              </table>
              
            </td>
          </tr>
          </tbody> 
        </table>
      </form>
      <table height=129 cellspacing=0 cellpadding=5 width=330 align=center 
      border=0>
        <tbody> 
        <tr> 
          <td align=left width=365 background=images/freereg_bg.gif 
          height=119><span class=sing>��û��ע�᱾վ��Ա��</span><br>
            ������������ӣ�ֻҪ<img 
            height=20 src="images/1min.gif" 
            width=52>���ӵ�ʱ��ע���Ϊ��վ��Ա��������ȫ������������һͬ�����������⡣<br>
            <br>
            <center>
              <a href="../reg/User_Reg.asp"><img height=31 src="images/1resister.gif" width=98 
                  border=0></a> 
            </center>
          </td>
        </tr>
        </tbody> 
      </table>
      
    </TD>
  </TR>
</TABLE>
<!--#include file="../Inc/down.htm"-->
</body>
</HTML>