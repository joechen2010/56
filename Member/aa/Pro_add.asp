<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<title><%=WebName%>-��Ա����ƽ̨</title>
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
<script LANGUAGE="JavaScript">
function check()
{
if (document.Form1.cpmc.value=="")
{
alert("��Ʒ���Ʋ���Ϊ��")
document.Form1.cpmc.focus()
document.Form1.cpmc.select() 
return
}
if (document.Form1.cpbh.value=="")
{
alert("��Ʒ��Ų���Ϊ��")
document.Form1.cpbh.focus()
document.Form1.cpbh.select()
return
}
if (document.Form1.cpsb.value=="")
{
alert("��Ʒ�̱겻��Ϊ��")
document.Form1.cpsb.focus()
document.Form1.cpsb.select()
return
}
if (document.Form1.cpcd.value=="")
{
alert("��Ʒ���ز���Ϊ��")
document.Form1.cpcd.focus()
document.Form1.cpcd.select() 
return
}

if (document.Form1.jysm.value=="")
{
alert("��Ҫ˵������Ϊ�գ�")
document.Form1.jysm.focus()
document.Form1.jysm.select() 
return
}

document.Form1.submit() 
}
</SCRIPT>
</head>
<BODY text=#000000 leftMargin=0 topMargin=0 
marginheight="0" marginwidth="0" bgcolor="#F7FBFF">
<!--#include file="head.htm"-->
<table width="750"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">����������</a> &gt; ������Ʒ��Ϣ</td>
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
                        <div align="center"><font color="#FFFFFF">������Ʒ��Ϣ</font></div>
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
                              <td align="left">
                        <div id="TipAll" > <font color="#FF0000">ע���벻Ҫ�����ظ���Ʒ��лл����&nbsp;&nbsp;&nbsp; 
                          </font></div>
                      </td>
                            </tr>
                           </table>
                 
                  <%
dim MaxP,hydj
set rs_c=conn.execute("select count(*) from Spzs where gsid="&session("id"))
if session("hydj")=0 then
    MaxP=0
	hydj="��ѻ�Ա" 
elseif session("hydj")=1 then
    MaxP=15  
	hydj="ͭ�ƻ�Ա" 
elseif session("hydj")=2 then
    MaxP=30
	hydj="���ƻ�Ա" 
elseif session("hydj")=3 then
    MaxP=99999
	hydj="���ƻ�Ա"   
else
    MaxP=0  
end if
if rs_c(0)>=MaxP then
response.write "<p>�Բ�������"&hydj&"ֻ�ܷ���"&MaxP&"����Ϣ��</p>"
end if
%>
                  <table border="0" cellspacing="1" style="border-collapse: collapse" width="100%" id="AutoNumber2" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" cellpadding="5" bgcolor="#B1D4F2">
                    <form method="POST" action="Pro_save_add.asp" name="Form1">
                      <tr> 
                        <TD align=right width=273 height="38" bgcolor="#EDF6FF">��Ʒ����<SPAN class=word9>��</SPAN></td>
                        <TD height="38" bgcolor="#EDF6FF">
                          <input class="f11" name="cpmc" size="31">
                          &nbsp;<font color="#FF6600">*</font></td>
                      </tr>
                      <tr> 
                        <TD align="right" width="273" height="23" bgcolor="#EDF6FF">��Ʒ��ţ�</TD>
                        <TD width=197 height="23" bgcolor="#EDF6FF">
                          <input class="f11" name="cpbh" size="15" title=�����Ʒ��ţ���LH��002����Ų�Ҫ�ظ����������϶�����>
                          &nbsp;<font color="#FF6600">*</font></td>
                      </tr>
                      <tr> 
                        <TD align="right" width="273" height="27" bgcolor="#EDF6FF">��Ʒ���أ�</TD>
                        <TD width=197 height="27" bgcolor="#EDF6FF">
                          <input class="f11" name="cpcd" size="15" title=�����Ʒ��ţ���LH��002>
                        </td>
                      </tr>
                      <tr> 
                        <TD align="right" width="273" height="26" bgcolor="#EDF6FF">��Ʒ�̱�<span class=word9>��</span></td>
                        <TD height="26" bgcolor="#EDF6FF"> 
                          <input class="f11" name="cpsb" size="15" title=���磺ɣ���ɡ�������e��Ƽ� >
                          &nbsp;<font color="#FF6600">*</font></TD>
                      </tr>
                      <tr> 
                        <TD align="right" width="273" height="26" bgcolor="#EDF6FF">��ƷͼƬ��</td>
                        <TD height="26" bgcolor="#EDF6FF"> &nbsp;<iframe name="Picture" frameborder=0 width=430 height=20 scrolling=no src=Upload.asp></iframe> 
                          <br>
                          ͼƬ·���� 
                          <input name="picture" type="TEXT" id="picture"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100 readonly>
                        </TD>
                      </tr>
                      <tr> 
                        <TD align="right" width="273" height="52" bgcolor="#EDF6FF" valign="middle"> 
                          ��Ҫ˵����<br>
                          &nbsp;(���100��)</TD>
                        <TD align="left" height="52" bgcolor="#EDF6FF">
                          <TEXTAREA class="f11" rows="3" name="jysm" cols="51" style="font-family: ����" title=��Ҫ˵�����ֻ��100����></TEXTAREA>
                          <FONT color="#FF6600">&nbsp;*</FONT></TD>
                      </tr>
                      <tr> 
                        <TD colspan="2" align="center" height="12" bgcolor="#EDF6FF"> 
                          <input type="hidden" class="smallinput" name="jptj" value="yes">
                          <input class="f11" type="button" value="ȷ ��" onClick="check()" name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                          <input class="f11" type="button" value="�� ��" name="button">
                        </TD>
                      </tr>
                    </FORM>
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


