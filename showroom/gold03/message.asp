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
<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center>
  <TBODY> 
  <TR bgColor=#fc8410> 
    <TD colSpan=2 height=3></TD>
  </TR>
  <TR bgColor=#cf6903> 
    <TD colSpan=2 height=2></TD>
  </TR>
  <TR> 
    <TD vAlign=top align=middle width=156 bgColor=#ff822e>
<!--#include file="left.htm"-->      
	  </TD>
    <TD vAlign=top align=middle bgColor=#fff6f1> 
      <TABLE cellSpacing=0 cellPadding=0 width="75%">
        <TBODY> 
        <TR> 
          <TD height=8></TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE height=142 cellSpacing=0 cellPadding=0 width="97%">
        <TBODY> 
        <TR> 
          <TD vAlign=top align=middle> 
            <TABLE cellSpacing=0 cellPadding=0 width="100%">
              <TBODY> 
              <TR> 
                <TD height=32><IMG height=32 src="images/menule.gif" 
                  width=9></TD>
                <TD width="100%" background=images/menumi.gif> 
                  <TABLE cellSpacing=0 cellPadding=0 width="75%">
                    <TBODY> 
                    <TR> 
                      <TD width="5%"><IMG height=14 
                        src="images/iconrighl.gif" width=14></TD>
                      <TD class=font14 vAlign=bottom 
                    width="95%">��������</TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                </TD>
                <TD><IMG height=32 src="images/menuri.gif" 
              width=11></TD>
              </TR>
              </TBODY> 
            </TABLE>
            <TABLE class=btBorder cellSpacing=0 cellPadding=0 width="100%" 
            bgColor=#ffffff>
              <TBODY> 
              <TR> 
                <TD vAlign=top align=middle> 
                  <TABLE class=style cellSpacing=5 cellPadding=0 width="100%" 
                  align=center border=0>
                    <TBODY> 
                    <TR> 
                      <TD class=S2 align=middle bgColor=#f9f9f9></TD>
                    </TR>
                    <TR> 
                      <TD class=S> 
                        <TABLE height=24 cellSpacing=0 cellPadding=0 width="100%" 
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

if (document.Form1.F_Fax.value =="")
{
alert("����д������룡");
document.Form1.F_Fax.focus();
document.Form1.F_Fax.select();
return false;
}
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
                                          <td width="96" align="right"><font color="#000000">�ռ��ˣ�</font></td>
                                          <td width="503"><%=rs_q("name")%></td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">* 
                                            </font><font color="#000000"></font><font color="#000000">���⣺</font></td>
                                          <td width="503"> 
                                            <input type="text" name="subject" value="ѯ��ҵ��������ϵ">
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">* 
                                            </font>���ݣ�</td>
                                          <td width="503"> 
                                            <textarea cols=50 name=content rows=6 style="border-style: solid; border-width: 1"></textarea>
                                          </td>
                                        </tr>
                                        <tbody> 
                                        <tr> 
                                          <td width="96" align="right"><font color=#ff6600>*</font> 
                                            ���Ĺ�˾��</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Company" value=<%=session("qymc")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color=#ff6600>*</font> 
                                            �������֣�</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Name" value=<%=session("name")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color=#ff6600>*</font> 
                                            �����ʼ���</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Email" value=<%=session("email")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color=#ff6600>*</font> 
                                            ��ϵ��ַ��</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_address" value=<%=session("address")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color=#ff6600>*</font> 
                                            ��&nbsp;&nbsp;&nbsp;&nbsp;��:</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Post" value=<%=session("post")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"> <font color=#ff6600>*</font> 
                                            ��&nbsp;&nbsp;&nbsp;&nbsp;����</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Tel" value=<%=session("phone")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right">&nbsp; 
                                            <font color=#ff6600>*</font> ��&nbsp;&nbsp;&nbsp;&nbsp;��:</td>
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
                                            <img height=1 width=80> 
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
                      </TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                  </TD>
              </TR>
              </TBODY> 
            </TABLE>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
      
    </TD>
  </TR>
  </TBODY> 
</TABLE>

<!--#include file="Copyright.htm"--></BODY></HTML>
<%end if%>
