<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"--> 
<!--#include file="check.asp"-->
<%
set rs_q=conn.execute("select * from qyml where user='"+session("user")+"'")
if rs_q.eof and rs_q.bof then
    response.write ("�Բ��𣬲�������<a href=""javascript:history.back(-1)"">�뷵��</a>")
	response.end
else
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
<%
set rs=server.createobject("adodb.recordset")
%>
<script Language="JavaScript"> 
function FormCheck() 
{ 
if (document.Form1.question.value =="") 
{ 
alert("������������ʾ���⣡"); 
document.Form1.question.focus(); 
return (false); 
} 
if (document.Form1.answer.value =="") 
{ 
alert("������������ʾ�𰸣�"); 
document.Form1.answer.focus(); 
return (false); 
} 
if (document.Form1.name.value =="") 
{ 
alert("����������������"); 
document.Form1.name.focus(); 
return (false); 
} 
if (document.Form1.province.value =="ʡ��") 
{ 
alert("��ѡ��ʡ�ݣ�"); 
document.Form1.province.focus(); 
return (false); 
} 
if (document.Form1.city.value =="") 
{ 
alert("��ѡ���������"); 
document.Form1.city.focus(); 
return (false); 
} 
if (document.Form1.address.value =="") 
{ 
alert("������������ϵ��ַ��"); 
document.Form1.address.focus(); 
return (false); 
} 
if (document.Form1.post.value =="") 
{ 
alert("�������������룡"); 
document.Form1.post.focus(); 
return (false); 
} 
if (document.Form1.email.value =="")
{
alert("���������ĵ����ʼ���ַ��");
document.Form1.email.focus();
document.Form1.email.select();
return false;
}
var filter=/^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
if (!filter.test(document.Form1.email.value)) { 
alert("�ʼ���ַ����ȷ,��������д��"); 
document.Form1.email.focus();
document.Form1.email.select();
return false; 
} 
document.Form1.submit()
} 
</script>
<script language=JavaScript src="../inc/p_c_a.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" bgcolor="#FFFFFF"  onload="setup()">
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
            <a href="http://www.cx56w.com">����������</a> &gt; �޸���ϵ��Ϣ</td>
          <td width="17%"><img src="images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="534"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> <%=session("user")%></font> </strong>����ӭ��������</td>
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
                  <TABLE border=0 
width="100%" height="225" style="border-collapse: collapse" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" cellpadding="5" cellspacing="1" bgcolor="#B1D4F2">
                    <FORM method="POST" action="modify_info_save.asp" name="Form1" onsubmit="return FormCheck();" >
                      <tr> 
                        <TD align=left height="25" colspan="2" width="396" class="white"  background="images/title_bg.gif">�� 
                          �ʻ��������� </TD>
                      </TR>
                      <tr> 
                        <TD align=right height="22" width="780" colspan="2" bgcolor="#D2EAFF"> 
                          <p align="left"> 
                        </TD>
                      </tr>
                      <TR valign="middle"> 
                        <TD align=right width="150" height="38" bgcolor="#D2EAFF">��Ա�ʺţ�</TD>
                        <TD height="38" width="640" bgcolor="#EEF7FF"> 
                          <INPUT class="box" maxLength=20 name="user" size="20" value="<%=rs_q("user")%>" readonly>
                          <font color="#FF0000"> *</font><font color="#666666">&nbsp;4-20���ַ���<font face="Arial">A-Z</font>�� 
                          <font face="Arial"> a-z</font>�� <font face="Arial"> 
                          0-9</font>��</font></TD>
                      </TR>
                      <TR vAlign="middle"> 
                        <TD align=right height="39" width="165" bgcolor="#D2EAFF">������ʾ��</TD>
                        <TD height="39" width="640" bgcolor="#EEF7FF"> 
                          <INPUT class="box" maxLength=20 name=question value="<%=rs_q("question")%>" size="20">
                          <FONT color=#990000>&nbsp; </FONT> <FONT color=#FF0000>*</FONT><font color="#666666">�������������룬��Ҫ�һ������ʱ�����ǵ�ϵͳ��Ҫ�����ش�������⡣</font></TD>
                      </TR>
                      <TR vAlign="middle"> 
                        <TD align=right height="50" width="165" bgcolor="#D2EAFF">��ʾ�𰸣�</TD>
                        <TD height="50" width="640" bgcolor="#EEF7FF"> 
                          <INPUT class="box" maxLength=20 name=answer value="<%=rs_q("answer")%>" size="20">
                          <FONT color=#FF0000 
size=2>*</FONT><font color="#666666">������ĵ绰������85694771�������Ϳ��԰�������ʾ����дΪ��85694771��&nbsp;�����μ�����𰸣��Ա����붪ʧʱ�ش�ϵͳ�����ʡ�������Ļش���ȷ��ϵͳ�ͻ��Զ���������ʾ��������</font></TD>
                      </TR>
                      <tr> 
                        <TD height="34" colspan="2" class="white" background="images/title_bg.gif">�� �޸Ļ�������</TD>
                      </tr>
                      <tr> 
                        <TD align=right height="34" width="165" bgcolor="#D2EAFF">����������</TD>
                        <TD height="34" width="262" bgcolor="#EEF7FF"> 
                          <INPUT class="box" maxLength=100 name=name size="20" value=<%=rs_q("name")%>>
                          <FONT color=#FF0000 size=2>*</FONT></TD>
                      </tr>
                      <tr> 
                        <TD align=right height="18" width="115" bgcolor="#D2EAFF">��&nbsp; 
                          ��</TD>
                        <TD height="18" width="291" bgcolor="#EEF7FF"> 
                          <input name=ch type=radio <%if rs_q("ch") ="����" then Response.Write "checked"%> value="����" checked>
                          ����&nbsp;&nbsp; 
                          <input name=ch type=radio <%if rs_q("ch") ="Ůʿ" then Response.Write "checked"%> value="Ůʿ" size="20">
                          Ůʿ</TD>
                      </tr>
                      <tr> 
                        <TD align=right height="18" width="115" bgcolor="#D2EAFF"> 
                          <p align="right">�������� 
                        </TD>
                        <TD height="18" width="291" bgcolor="#EEF7FF"> 
              <select id="s1" name="province">
                <option value="ʡ��">ʡ��</option>
              </select>
              <select id="s2" name="city">
                <option>�ؼ���</option>
              </select>
              <select id="s3" name="area">
                <option>�С��ؼ��С���</option>
              </select>           
                        </TD>
                      </tr>
                      <tr> 
                        <TD align=right height="33" width="165" bgcolor="#D2EAFF">��ϵ��ַ��</TD>
                        <TD height="33" width="262" bgcolor="#EEF7FF"> 
                          <INPUT class="box" maxLength=250 name=address size=29 value=<%=rs_q("address")%>>
                          <FONT color=#FF0000 size=2>*</FONT></TD>
                      </tr>
                      <tr> 
                        <TD align=right height="33" width="165" bgcolor="#D2EAFF">��&nbsp; 
                          �ࣺ</TD>
                        <TD height="33" width="262" bgcolor="#EEF7FF"> 
                          <input class="box" maxlength=8 name=post size=8 value=<%=rs_q("post")%>>
                          <font color=#FF0000 size=2>*</font></TD>
                      </tr>
                      <tr> 
                        <TD align=right height="34" width="165" bgcolor="#D2EAFF"> 
                          <p align="right">����ְ�� 
                        </TD>
                        <TD height="34" width="262" bgcolor="#EEF7FF"> 
                          <INPUT class="box" name=zw size="20" value=<%=rs_q("zw")%>>
                        </TD>
                      </tr>
                      <tr> 
                        <TD align=right height="34" width="165" bgcolor="#D2EAFF">��&nbsp; 
                          ����</TD>
                        <TD height="34" width="262" bgcolor="#EEF7FF"> 
                          <input class="box" name=phone1 size=5 value="<%=rs_q("phone1")%>">-<input class="box" name=phone2 size=5 value="<%=rs_q("phone2")%>">-<input class="box" name=phone3 size=10 value="<%=rs_q("phone3")%>">
                          <font color=#FF0000 size=2>*</font></TD>
                      </tr>
                      <tr> 
                        <TD align=right height="34" width="165" bgcolor="#D2EAFF">��&nbsp;&nbsp;�棺</TD>
                        <TD height="34" width="262" bgcolor="#EEF7FF"> 
                          <input class="box" name=fax1 size=5 value=<%=rs_q("fax1")%>>-<input class="box" name=fax2 size=5 value=<%=rs_q("fax2")%>>-<input class="box" name=fax3 size=10 value=<%=rs_q("fax3")%>>
                          <font color=#FF0000 size=2>*</font>
                        </TD>
                      </tr>
                      <tr> 
                        <TD align=right height="34" width="165" bgcolor="#D2EAFF">�ֻ����룺</TD>
                        <TD height="34" width="262" bgcolor="#EEF7FF"> 
                          <INPUT class="box" name=mobile size=20 value="<%=rs_q("mobile")%>">
                        </TD>
                      </tr>
                      <tr> 
                        <TD align=right height="36" width="165" bgcolor="#D2EAFF">�����ʼ���</TD>
                        <TD height="36" width="262" bgcolor="#EEF7FF"> 
                          <INPUT class="box" maxLength=100 name=email size=29 value=<%=rs_q("email")%>>
                          <FONT color=#FF0000 size=2>*</FONT>&nbsp;</TD>
                      </tr>
                      <tr> 
                        <TD align=right height="36" width="165" bgcolor="#D2EAFF">��&nbsp;&nbsp;ַ��</TD>
                        <TD height="36" width="262" bgcolor="#EEF7FF"> 
                          <input class="box" maxlength=100 name=web size=25 value=<%=rs_q("web")%>>
                        </TD>
                      </tr>
                      <tr> 
                        <TD align=right height="1" width="780" colspan="2" bgcolor="#EEF7FF"></TD>
                      </tr>
                      <tr> 
                        <TD align=right height="42" width="780" colspan="2" bgcolor="#D2EAFF"> 
                          <p align="center"> 
                            <input type="submit" value=" �޸� "class="btsub">
                            &nbsp;&nbsp;&nbsp; 
                            <INPUT name=Reset type=reset value=" ���� " class="btsub">
                            <br>
                        </TD>
                      </tr>
                    </FORM>
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
  
</BODY></HTML>
<%
end if
rs_q.close
set rs_q=nothing
conn.close
set conn=nothing
%>