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
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 给我留言</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width="16%" background=images/bj.jpg> 
      <TABLE height=254 cellSpacing=0 cellPadding=0 width=250 
      background=images/bj.jpg border=0>
        <TBODY> 
        <TR> 
          <TD vAlign=top> 
            <TABLE height=23 cellSpacing=0 cellPadding=0 width="100%" 
              border=0>
              <TBODY> 
              <TR> 
                <TD class=zi2 background=images/syt_6.jpg 
                height=23></TD>
              </TR>
              </TBODY> 
            </TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY> 
              <TR> 
                <TD vAlign=bottom background=images/l2.jpg><img 
                  height=222 src="images/syt_8.jpg" width=251 usemap=#Map border=0>
                  <table cellspacing=0 cellpadding=0 width="100%" 
                  background=images/l2.jpg border=0>
                    <tbody> 
                    <tr> 
                      <td valign=top> 
                        <table cellspacing=2 cellpadding=8 width=222 
                        align=right border=0>
                          <tbody> 
                          <tr> 
                            <td align=middle height="15">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td align=middle> 
                              <table width="100%" border="0" cellpadding="0" cellspacing="0" class=S2 style="border-top:1px dotted #FF6600">
                                <tr> 
                                  <td width="22%" nowrap class=S2 height="28">地址：</td>
                                  <td width="78%" colspan="3" class=S2 height="28"><%=rs_q("address")%></td>
                                </tr>
                                <tr> 
                                  <td class=S2 width="22%" height="28">邮编：</td>
                                  <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"><%=rs_q("post")%></font></td>
                                </tr>
                                <tr> 
                                  <td class=S2 width="22%" height="28">电话：</td>
                                  <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"> 
                                    <% Response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3")%>
                                    </font></td>
                                </tr>
                                <tr> 
                                  <td class=S2 width="22%" height="28">传真：</td>
                                  <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"> 
                                    <% Response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
                                    </font></td>
                                </tr>
                                <tr> 
                                  <td class=S2 width="22%" height="28">主页：</td>
                                  <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a>&nbsp; 
                                    </font></td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                          </tbody> 
                        </table>
                        <table cellspacing=0 cellpadding=0 width="80%" 
                        align=center border=0>
                          <tbody> 
                          <tr> 
                            <td></td>
                          </tr>
                          </tbody> 
                        </table>
                      </td>
                    </tr>
                    <tr> 
                      <td><img height=18 src="images/l1.jpg" 
                      width=251></td>
                    </tr>
                    </tbody> 
                  </table>
                </TD>
              </TR>
              </TBODY> 
            </TABLE>
            <BR>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD>
    <TD vAlign=top width="84%"> 
      <TABLE height=80 cellSpacing=0 cellPadding=0 width="100%" 
      background=images/syt_7.jpg border=0>
        <TBODY> 
        <TR> 
          <TD width="86%"> 
            <DIV class=tile> 
              <DIV align=right><STRONG><%=rs_q("qymc")%></STRONG></DIV>
            </DIV>
          </TD>
          <TD width="14%"> 
            <DIV align=right> <OBJECT 
            codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0 
            height=80 width=200 
            classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000>
                <PARAM NAME="movie" VALUE="images/line1.swf">
                <PARAM NAME="quality" VALUE="high">
                <PARAM NAME="wmode" VALUE="transparent">
                <embed 
            src="images/line1.swf" quality="high" 
            pluginspage="http://www.macromedia.com/go/getflashplayer" 
            type="application/x-shockwave-flash" width="200" 
            height="80">
                </embed> 
              </OBJECT></DIV>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE height=23 cellSpacing=0 cellPadding=0 width="100%" bgColor=#fff4f4 
      border=0>
        <TBODY> 
        <TR> 
          <TD class=zi2 width="97%" height=23> 
            <DIV align=right><FONT color=#ff0000>物流通金牌会员：创建于<%=rs_q("idate")%></FONT></DIV>
          </TD>
          <TD class=zi2 width=10>&nbsp;</TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE height=96 cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY> 
        <TR> 
          <TD height=96><BR>
            <TABLE height=37 cellSpacing=0 cellPadding=0 width="90%" 
            align=center background=images/title.jpg border=0>
              <TBODY> 
              <TR> 
                <TD width="6%"><IMG height=32 src="images/pic.jpg" 
                  width=38></TD>
                <TD width="22%"> 
                  <TABLE height=24 cellSpacing=0 cellPadding=0 width="69%" 
                  align=center bgColor=#ffffff border=0>
                    <TBODY> 
                    <TR> 
                      <TD> 
                        <DIV class=tile2 
                      align=center><STRONG>给我留言</STRONG></DIV>
                      </TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                </TD>
                <TD width="72%">&nbsp;</TD>
              </TR>
              </TBODY> 
            </TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY> 
              <TR> 
                <TD><BR>
                  <TABLE class=style cellSpacing=5 cellPadding=0 width="95%" 
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
alert("请填写发送的内容！")
document.Form1.content.focus()
document.Form1.content.select()
return false;
}

if (document.Form1.F_Company.value=="")
{
alert("请填写公司名称")
document.Form1.F_Company.focus()
document.Form1.F_Company.select()
return false;
}
if (document.Form1.F_Name.value=="")
{
alert("请填写联系人的姓名！")
document.Form1.F_Name.focus()
document.Form1.F_Name.select()
return false;
}

if (document.Form1.F_Email.value =="")
{
alert("请输入您的电子邮件地址！");
document.Form1.F_Email.focus();
document.Form1.F_Email.select();
return false;
}
var filter=/^\s*([A-Za-z0-9_-]+(\.\w+)*@(\w+\.)+\w{2,3})\s*$/;
if (!filter.test(document.Form1.F_Email.value)) { 
alert("邮件地址不正确,请重新填写！"); 
document.Form1.F_Email.focus();
document.Form1.F_Email.select();
return false; 
}  

if (document.Form1.F_address.value=="")
{
alert("请填写详细地址")
document.Form1.F_address.focus()
document.Form1.F_address.select()
return false;
} 
if (document.Form1.F_Post.value=="")
{
alert("请填写邮政编码")
document.Form1.F_Post.focus()
document.Form1.F_Post.select()
return false;
} 

if (document.Form1.F_Tel.value =="")
{
alert("请输入您的电话号码！");
document.Form1.F_Tel.focus();
document.Form1.F_Tel.select();
return false;
}
if (document.Form1.F_Fax.value =="")
{
alert("请填写传真号码！");
document.Form1.F_Fax.focus();
document.Form1.F_Fax.select();
return false;
}
document.Form1.submit()
}
</SCRIPT>
                              　　　如果你对我公司有有任何意见和建议，请留下您的信息，我们将尽快与您联系。 
                              <form method="POST" action="messagesave.asp" name="Form1"  onSubmit="return CheckForm();">
                                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber10">
                                  <tr> 
                                    <td width="100%" height="267"> 
                                      <table border=0 cellpadding=2 cellspacing=2 width="98%" align="center">
                                        <tr> 
                                          <td width="96" align="right">收件人公司：</td>
                                          <td width="503"><b class="L"><%=rs_q("qymc")%></b></td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#000000">收 
                                            件 人：</font></td>
                                          <td width="503"><%=rs_q("name")%></td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">* 
                                            </font><font color="#000000"></font><font color="#000000">主　　题：</font></td>
                                          <td width="503"> 
                                            <input type="text" name="subject" value="询问业务，请速联系">
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">* 
                                            </font>内　　容：</td>
                                          <td width="503"> 
                                            <textarea cols=50 name=content rows=6 style="border-style: solid; border-width: 1"></textarea>
                                          </td>
                                        </tr>
                                        <tbody> 
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">*</font> 
                                            您的公司：</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Company" value=<%=session("qymc")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">*</font> 
                                            您的名字：</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Name" value=<%=session("name")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">*</font> 
                                            您的邮件：</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Email" value=<%=session("email")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">*</font> 
                                            联系地址：</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_address" value=<%=session("address")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"><font color="#FF0000">*</font> 
                                            邮　　编：</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Post" value=<%=session("post")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"> <font color="#FF0000">*</font> 
                                            电　　话：</td>
                                          <td width="503"> 
                                            <input type="text" size="33" name="F_Tel" value=<%=session("phone")%>>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="96" align="right"> <font color="#FF0000">*</font> 
                                            传　　真：</td>
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
                                            <input type="submit" value="确 定" name="button" onclick="">
                                            <img height=1 width=80> 
                                            <input name=reset class="smallinput" type="button" value="返 回" onClick=javascript:history.go(-1)>
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
  </TR></TBODY></TABLE>
<!--#include file="Copyright.htm"-->
                  <map name="Map"> 
                    <area shape="rect" coords="110,63,221,89" href="../../bearingpass" target="_blank">
                    <area shape="rect" coords="110,99,221,125" href="../../About/update_esc.asp" target="_blank">
                  </map>
				  </BODY></HTML>
<%end if%>
