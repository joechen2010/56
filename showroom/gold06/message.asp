<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/function.asp"-->
<%
Dim Gsid
Gsid=SafeRequest("Gsid",1)
if gsID="" then
     call WriteErrMsg2()
end if
set rs_q=server.createobject("adodb.recordset")
sql="select * from qyml where id="&gsid&""
rs_q.open sql,conn,1,1
if rs_q.eof and rs_q.bof then
     call WriteErrMsg2()
else
%>
<HTML>
<HEAD>
<TITLE><%=rs_q("qymc")%> - ����ͨ������ҵ - ��������</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
<style>
a{
	font-size: 10pt;
	color: #000000;
	text-decoration: none;
	FONT-FAMILY: "����"
}
a:hover {
	color: #FF0000;
	text-decoration: none;
}
td {
	font-size: 10pt;
	text-decoration: blink;
}
</style>
</HEAD>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="head.htm" -->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="169" background="images/mb_01.jpg"><DIV class=tile> ����<font color="#FF0000" size="+3"><STRONG><%=rs_q("qymc")%></STRONG></font>
            </DIV></td>
  </tr>
</table>
<table width="1003" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="200" align="center" valign="top">
<!--#include file="left.htm" -->
      <table width="160" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="20">&nbsp;</td>
          <td width="120">&nbsp;</td>
          <td width="20">&nbsp;</td>
        </tr>
        <tr> 
          <td background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td background="images/bk_top2.jpg"> <DIV class=tile2 
              align=center><STRONG> ��ϵ����</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr> 
          <td  width="160" colspan="3" valign="top" nowrap class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999; border-top: 1px solid #999999;">&nbsp;</td>
        </tr>
        <tr bgcolor="#F6F6F6"> 
          <td  width="160" height="25" colspan="3" nowrap class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">��ַ��<%=rs_q("address")%></td>
        </tr>
        <tr> 
          <td width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;" td>�ʱࣺ<font face="Arial"><%=rs_q("post")%></font></td>
        </tr>
        <tr bgcolor="#F6F6F6"> 
          <td width="160" height="25" colspan="3"  class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">�绰��<font face="Arial"> 
            <% Response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3")%>
            </font></td>
        </tr>
        <tr> 
          <td width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">���棺<font face="Arial"> 
            <% Response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
            </font></td>
        </tr>
        <tr bgcolor="#F6F6F6"> 
          <td width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;"><font face="Arial">E-mail��</font><font face="Arial"><a href="mailto:<%=rs_q("email")%>"><%=rs_q("email")%></a></font></td>
        </tr>
        <tr> 
          <td width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;border-bottom: 1px solid #999999">��ҳ��<font face="Arial"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a></font></td>
        </tr>
      </table>
    </td>
    <td width="803" align="center" valign="top"> <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td width="731" background="images/bk_top2.jpg"> <DIV class=tile2 
              ><STRONG>��������</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr align="center"> 
          <td colspan="3"  class=C vAlign=top borderColor=#cccccc style="border: 1px solid #999999;"><br>
            <TABLE height=24 cellSpacing=0 cellPadding=0 width="98%" 
border=0 style="border-collapse: collapse" bordercolor="#111111">
              <TBODY> 
                          <TR> 
                            <TD vAlign=top width="29%" height=148> 
                              <script LANGUAGE="JavaScript">
function CheckForm()
{
if (document.Form1.content.value=="")
{
alert("����д���͵����ݣ�")
document.Form1.content.focus()
document.Form1.content.select()
return false;
}

if (document.Form1.F_Company.value=="")
{
alert("����д��˾����")
document.Form1.F_Company.focus()
document.Form1.F_Company.select()
return false;
}
if (document.Form1.F_Name.value=="")
{
alert("����д��ϵ�˵�������")
document.Form1.F_Name.focus()
document.Form1.F_Name.select()
return false;
}

if (document.Form1.F_Email.value =="")
{
alert("���������ĵ����ʼ���ַ��");
document.Form1.F_Email.focus();
document.Form1.F_Email.select();
return false;
}
var filter=/^\s*([A-Za-z0-9_-]+(\.\w+)*@(\w+\.)+\w{2,3})\s*$/;
if (!filter.test(document.Form1.F_Email.value)) { 
alert("�ʼ���ַ����ȷ,��������д��"); 
document.Form1.F_Email.focus();
document.Form1.F_Email.select();
return false; 
}  

if (document.Form1.F_address.value=="")
{
alert("����д��ϸ��ַ")
document.Form1.F_address.focus()
document.Form1.F_address.select()
return false;
} 
if (document.Form1.F_Post.value=="")
{
alert("����д��������")
document.Form1.F_Post.focus()
document.Form1.F_Post.select()
return false;
} 

if (document.Form1.F_Tel.value =="")
{
alert("���������ĵ绰���룡");
document.Form1.F_Tel.focus();
document.Form1.F_Tel.select();
return false;
}
var reg=/(^[0-9]{3,4}\-[0-9]{3,8}$)|(^[0-9]{3,8}$)|(^\([0-9]{3,4}\)[0-9]{3,8}$)|(^0{0,1}13[0-9]{9}$)/
if (!reg.test(document.Form1.F_Tel.value)) { 
alert("���������ĵ绰���룡"); 
document.Form1.F_Tel.focus();
document.Form1.F_Tel.select();
return false; 
}

if (document.Form1.F_Fax.value =="")
{
alert("����д������룡");
document.Form1.F_Fax.focus();
document.Form1.F_Fax.select();
return false;
}
var reg=/(^[0-9]{3,4}\-[0-9]{3,8}$)|(^[0-9]{3,8}$)|(^\([0-9]{3,4}\)[0-9]{3,8}$)|(^0{0,1}13[0-9]{9}$)/
if (!reg.test(document.Form1.F_Fax.value)) { 
alert("����д������룡"); 
document.Form1.F_Fax.focus();
document.Form1.F_Fax.select();
return false; 
}
document.Form1.submit()
}
</SCRIPT>
                              �������ҹ�˾�����κ�����ͽ��飬������������Ϣ�����ǽ�����������ϵ�� 
                              <form method="POST" action="messagesave.asp" name="Form1"  onSubmit="return CheckForm();">
                                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber10">
                                  <tr> 
                                    <td width="100%" height="267"> 
                                      <table border=0 cellpadding=2 cellspacing=2 width="98%" align="center">
                                        <tr> 
                                          <td width="96" align="right">�ռ��˹�˾��</td>
                                          <td width="503"><b class="L"><%=rs_q("qymc")%></b></td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#000000">�� 
                                            �� �ˣ�</font></td>
                                          <td width="503"><%=rs_q("name")%></td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">* 
                                            </font><font color="#000000">&nbsp;</font><font color="#000000">�������⣺</font></td>
                                          <td width="503"> 
                                            <input type="text" name="subject" value="ѯ��ҵ��������ϵ">
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">* 
                                            </font>�ڡ����ݣ�</td>
                                          <td width="503"> 
                                            <textarea cols=50 name=content rows=6 style="border-style: solid; border-width: 1"></textarea>
                                          </td>
                                        </tr>
                                        <tbody> 
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">*</font> 
                                            ���Ĺ�˾��</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Company" value=<%=session("qymc")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">*</font> 
                                            �������֣�</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Name" value=<%=session("name")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">*</font> 
                                            �����ʼ���</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Email" value=<%=session("email")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">*</font> 
                                            ��ϵ��ַ��</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_address" value=<%=session("address")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">*</font> 
                                            �ʡ����ࣺ</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Post" value=<%=session("post")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"> <font color="#FF0000">*</font> 
                                            �硡������</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Tel" value=<%=session("phone")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"> <font color="#FF0000">*</font> 
                                            �������棺</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Fax" value=<%=session("fax")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td colspan="2"> 
                                            <input type="hidden" name="F_User" value="<%=session("user")%>">
                                            <input type="hidden" name="T_User" value="<%=rs_q("user")%>">
                                            <input type="hidden" name="T_Name" value="<%=rs_q("name")%>">
                                            <input type="hidden" name="T_Company" value="<%=rs_q("qymc")%>">
                                            <input type="hidden" name="T_Email" value="<%=rs_q("email")%>">
                                          </td>
                                        </tr>
                                        <tr align="CENTER"> 
                                          <td colspan="2"> 
                                            <input type="submit" value="ȷ ��" name="button" onclick="">
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                                    <input name=reset class="smallinput" type="button" value="�� ��" onClick=javascript:history.go(-1)>
                                          </td>
                                        </tr>
                                        </tbody> 
                                      </table>
                                    </td>
                                  </tr>
                                </table>
                              </form>
                            </TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                  
		  </td>
        </tr>
        <tr> 
          <td height="30" colspan="3">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="Copyright.htm" -->
<%end if%>
</body>
</html>