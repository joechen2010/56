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
<LINK href="images/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<table cellspacing=0 cellpadding=0 width="100%" border=0>
  <tbody> 
  <tr> 
    <td valign=top width=189 bgcolor=#fcf6e8> 
      <table height=69 cellspacing=0 cellpadding=0 width="100%" 
      background=images/img_bg_9.gif border=0>
        <tbody> 
        <tr> 
          <td valign="bottom"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="20"> 
                  <div align="center"><font color=#ff0000>����ͨ��Ա��������<%=rs_q("idate")%></font></div>
                </td>
              </tr>
            </table>
          </td>
          <td align=right width=10><img height=69 src="images/img_11.gif" 
            width=10></td>
        </tr>
        </tbody> 
      </table>
      <!--#include file="left.htm"-->
      <table cellspacing=0 cellpadding=0 width=100% border=0>
        <tbody> 
        <tr> 
          <td align=middle height="15">&nbsp;</td>
        </tr>
        <tr> 
          <td align=middle> 
            <table width="96%" border="0" cellpadding="0" cellspacing="0" class=S2 style="border-top:1px dotted #FF6600" align="center">
              <tr> 
                <td width="22%" nowrap class=S2 height="28">��ַ��</td>
                <td width="78%" colspan="3" class=S2 height="28"><%=rs_q("address")%></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">�ʱࣺ</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"><%=rs_q("post")%></font></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">�绰��</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"> 
                  <% Response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3")%>
                  </font></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">���棺</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"> 
                  <% Response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
                  </font></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">��ҳ��</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a>&nbsp; 
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td valign=top align=middle> 
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
        <tbody> 
        <tr> 
          <td valign=top height=12><img height=10 src="images/img_12.gif" 
            width=11></td>
          <td valign=top align=right><img height=10 src="images/img_14.gif" 
            width=10></td>
        </tr>
        </tbody> 
      </table>
      <table cellspacing=0 cellpadding=0 width="96%" 
      background=images/img_bg_31.gif border=0>
        <tbody> 
        <tr> 
          <td width="4%" height=29><img height=29 src="images/img_30.gif" 
            width=13></td>
          <td class=font14 width="92%"><strong>��������</strong></td>
          <td align=right width="4%"><img height=29 src="images/img_33.gif" 
            width=10></td>
        </tr>
        </tbody> 
      </table>
      <table class=style cellspacing=5 cellpadding=0 width="95%" 
                  align=center border=0>
        <tbody> 
        <tr> 
          <td class=S2 align=middle bgcolor=#f9f9f9></td>
        </tr>
        <tr> 
          <td class=S> 
            <table height=24 cellspacing=0 cellpadding=0 width="100%" 
border=0 style="border-collapse: collapse" bordercolor="#111111">
              <tbody> 
              <tr> 
                <td valign=top width="29%" height=148> 
                  <script language="JavaScript">
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

if (document.Form1.F_Fax.value =="")
{
alert("����д������룡");
document.Form1.F_Fax.focus();
document.Form1.F_Fax.select();
return false;
}
}
</script>
                  �������������ҹ�˾�����κ�����ͽ��飬������������Ϣ�����ǽ�����������ϵ�� 
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
                                </font><font color="#000000"></font><font color="#000000">�������⣺</font></td>
                              <td width="503"> 
                                <input type="text" name="subject2" value="ѯ��ҵ��������ϵ">
                              </td>
                            </tr>
                            <tr> 
                              <td width="96" align="right"><font color="#FF0000">* 
                                </font>�ڡ����ݣ�</td>
                              <td width="503"> 
                                <textarea cols=56 name=textarea rows=6 style="border-style: solid; border-width: 1"></textarea>
                              </td>
                            </tr>
                            <tbody> 
                            <tr> 
                              <td width="96" align="right"><font color="#FF0000">*</font> 
                                ���Ĺ�˾��</td>
                              <td width="503"> 
                                <input type="text" size="40" name="F_Company2" value=<%=session("qymc")%>>
                              </td>
                            </tr>
                            <tr> 
                              <td width="96" align="right"><font color="#FF0000">*</font> 
                                �������֣�</td>
                              <td width="503"> 
                                <input type="text" size="40" name="F_Name2" value=<%=session("name")%>>
                              </td>
                            </tr>
                            <tr> 
                              <td width="96" align="right"><font color="#FF0000">*</font> 
                                �����ʼ���</td>
                              <td width="503"> 
                                <input type="text" size="40" name="F_Email2" value=<%=session("email")%>>
                              </td>
                            </tr>
                            <tr> 
                              <td width="96" align="right"><font color="#FF0000">*</font> 
                                ��ϵ��ַ��</td>
                              <td width="503"> 
                                <input type="text" size="40" name="F_address2" value=<%=session("address")%>>
                              </td>
                            </tr>
                            <tr> 
                              <td width="96" align="right"><font color="#FF0000">*</font> 
                                �ʡ����ࣺ</td>
                              <td width="503"> 
                                <input type="text" size="40" name="F_Post2" value=<%=session("post")%>>
                              </td>
                            </tr>
                            <tr> 
                              <td width="96" align="right"> <font color="#FF0000">*</font> 
                                �硡������</td>
                              <td width="503"> 
                                <input type="text" size="40" name="F_Tel2" value=<%=session("phone")%>>
                              </td>
                            </tr>
                            <tr> 
                              <td width="96" align="right"> <font color="#FF0000">*</font> 
                                �������棺</td>
                              <td width="503"> 
                                <input type="text" size="40" name="F_Fax2" value=<%=session("fax")%>>
                              </td>
                            </tr>
                            <tr> 
                              <td colspan="2"> 
                                <input type="hidden" name="F_User2" value="<%=session("user")%>">
                                <input type="hidden" name="T_User2" value="<%=rs_q("user")%>">
                                <input type="hidden" name="T_Name2" value="<%=rs_q("name")%>">
                                <input type="hidden" name="T_Company2" value="<%=rs_q("qymc")%>">
                                <input type="hidden" name="T_Email2" value="<%=rs_q("email")%>">
                              </td>
                            </tr>
                            <tr align="CENTER"> 
                              <td colspan="2"> 
                                <input type="submit" value="ȷ ��" name="button2" onClick="">
                                <img height=1 width=80> 
                                <input name=reset2 class="smallinput" type="button" value="�� ��" onClick=javascript:history.go(-1)>
                              </td>
                            </tr>
                            </tbody> 
                          </table>
                        </td>
                      </tr>
                    </table>
                  </form>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td valign=top width=15 bgcolor=#fcf6e8> 
      <table height=69 cellspacing=0 cellpadding=0 width="100%" 
      background=images/img_bg_9.gif border=0>
        <tbody> 
        <tr> 
          <td>&nbsp;</td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<!--#include file="Copyright.htm"-->
</BODY></HTML>
<%end if%>
