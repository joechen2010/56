<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
Dim Gsid
Gsid=SafeRequest("Gsid",1)
if gsID="" then
    call WriteErrMsg()
end if
set rs_q=server.createobject("adodb.recordset")
sql="select * from qyml where id="&gsid&""
rs_q.open sql,conn,1,1
if rs_q.eof and rs_q.bof then
    call WriteErrMsg()
else
%>
<HTML>
<HEAD>
<TITLE><%=rs_q("qymc")%> - 普通会员 - 给我留言</TITLE>
<STYLE type=text/css>BODY {
	FONT-SIZE: 5pt
}
TD {
	FONT-SIZE: 12px; COLOR: #174b03; FONT-FAMILY: 宋体
}
.f14 {
	FONT-SIZE: 14px; LINE-HEIGHT: 120%
}
.noline:link {
	COLOR: #000; TEXT-DECORATION: none
}
.noline:visited {
	COLOR: #000; TEXT-DECORATION: none
}
.noline:hover {
	COLOR: #f60; TEXT-DECORATION: none
}
</STYLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<SCRIPT language=JavaScript type=text/JavaScript>
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</SCRIPT>
<META content="MSHTML 6.00.2462.0" name=GENERATOR></HEAD>
<BODY oncontextmenu="return false" onselectstart="return false" 
ondragstart="return false" vLink=#174b03 aLink=#174b03 link=#174b03 topMargin=0 
onload="MM_preloadImages('images/MB_18.gif','images/MB_20.gif','images/MB_22.gif','images/MB_24.gif')">
<!--#include file="head.htm"-->
<TABLE height=373 cellSpacing=0 cellPadding=0 width=778 align=center 
bgColor=#ffffff border=0>
  <TBODY> 
  <TR> 
    <TD colSpan=2 height=30> 
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY> 
        <TR> 
          <TD width=168><IMG height=30 src="images/MB_12.gif" width=168></TD>
          <TH align=middle width=586 background=images/MB_17.gif><FONT 
            color=#ffffff size=3><%=rs_q("qymc")%></FONT></TH>
          <TD align=right width=24 background=images/MB_17.gif><IMG height=30 src="images/MB_19.gif" width=12></TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD>
  </TR>
  <TR> 
    <TD vAlign=top width=167 background=images/MB_38.gif><IMG height=123 src="images/MB_21.gif" width=167> 
      <A onmouseover="MM_swapImage('Image16','','images/MB_18.gif',1)" onmouseout=MM_swapImgRestore()       href="aboutus.asp?gsid=<%=gsid%>"><IMG height=34 src="images/MBh_18.gif" width=167 border=0 name=Image16></A> 
      <A onmouseover="MM_swapImage('Image17','','images/MB_20.gif',1)" onmouseout=MM_swapImgRestore() 
      href="trade.asp?gsid=<%=gsid%>"><IMG height=34 src="images/MBh_20.gif" width=167 border=0 name=Image17></A> 
      <A href="message.asp?gsid=<%=gsid%>"><IMG height=34 src="images/MB_22.gif" width=167 border=0></A> 
      <A onmouseover="MM_swapImage('Image19','','images/MB_24.gif',1)" onmouseout=MM_swapImgRestore() 
      href="contactus.asp?gsid=<%=gsid%>"><IMG height=34 src="images/MBh_24.gif" width=167 border=0 name=Image19></A> 
    </TD>
    <TD vAlign=top width=611 rowSpan=2> 
      <TABLE cellSpacing=0 cellPadding=0 width="83%" align=right border=0>
        <TBODY> 
        <TR> 
          <TD>&nbsp;</TD>
        </TR>
        <TR> 
          <TD><IMG height=12 src="images/MB_26.gif" width=601></TD>
        </TR>
        <TR> 
          <TD background=images/MB_28.gif height=35> 
            <TABLE cellSpacing=0 cellPadding=0 width="95%" align=center border=0>
              <TBODY> 
              <TR> 
                <TD><IMG height=26 src="images/syjh.gif" width=341></TD>
              </TR>
              <TR> 
                <TD height=10></TD>
              </TR>
              </TBODY> 
            </TABLE>
          </TD>
        </TR>
        <TR> 
          <TD background=images/MB_28.gif height=12> <br>
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
var reg=/(^[0-9]{3,4}\-[0-9]{3,8}$)|(^[0-9]{3,8}$)|(^\([0-9]{3,4}\)[0-9]{3,8}$)|(^0{0,1}13[0-9]{9}$)/
if (!reg.test(document.Form1.F_Tel.value)) { 
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
var reg=/(^[0-9]{3,4}\-[0-9]{3,8}$)|(^[0-9]{3,8}$)|(^\([0-9]{3,4}\)[0-9]{3,8}$)|(^0{0,1}13[0-9]{9}$)/
if (!reg.test(document.Form1.F_Fax.value)) { 
alert("请填写传真号码！"); 
document.Form1.F_Fax.focus();
document.Form1.F_Fax.select();
return false; 
}
document.Form1.submit()
}
</script>
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
                                    <td width="96" align="right">收 件 人：</td>
                                    <td width="503"><%=rs_q("name")%></td>
                                  </tr>
                                  <tr> 
                                    <td width="96" align="right"><font color="#FF0000">* 
                                      </font><font color="#000000"></font>主　　题：</td>
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
                                    <td width="96" align="right"><font color=#ff6600>*</font> 
                                      您的公司：</td>
                                    <td width="503"> 
                                      <input type="text" size="33" name="F_Company" value=<%=session("qymc")%>>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="96" align="right"><font color=#ff6600>*</font> 
                                      您的名字：</td>
                                    <td width="503"> 
                                      <input type="text" size="33" name="F_Name" value=<%=session("name")%>>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="96" align="right"><font color=#ff6600>*</font> 
                                      您的邮件：</td>
                                    <td width="503"> 
                                      <input type="text" size="33" name="F_Email" value=<%=session("email")%>>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="96" align="right"><font color=#ff6600>*</font> 
                                      联系地址：</td>
                                    <td width="503"> 
                                      <input type="text" size="33" name="F_address" value=<%=session("address")%>>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="96" align="right"><font color=#ff6600>*</font> 
                                      邮　　编：</td>
                                    <td width="503"> 
                                      <input type="text" size="33" name="F_Post" value=<%=session("post")%>>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="96" align="right"> <font color=#ff6600>*</font> 
                                      电　　话：</td>
                                    <td width="503"> 
                                      <input type="text" size="33" name="F_Tel" value=<%=session("phone")%>>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="96" align="right"> <font color=#ff6600>*</font> 
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
                                      <input type="submit" value="确 定" name="button" onClick="">
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
                      </td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </TD>
        </TR>
        <TR> 
          <TD background=images/MB_28.gif height=13>
            <table cellspacing=0 cellpadding=0 width="95%" align=center border=0>
              <tbody> 
              <tr> 
                <td><img height=26 src="images/lxwm.gif" width=345></td>
              </tr>
              <tr> 
                <td height=10></td>
              </tr>
              </tbody> 
            </table>
          </TD>
        </TR>
        <TR> 
          <TD background=images/MB_28.gif height=25> <BR>
            <table width="90%" border="0" cellpadding="5" cellspacing="0" class=S2 style="border-top:1px dotted #FF6600" align="center">
              <tr bgcolor="#f6f6f6"> 
                <td width="51" valign="top" nowrap class=S2>地　址：</td>
                <td width="482" colspan="3" class=S2><%=rs_q("address")%></td>
              </tr>
              <tr> 
                <td class=S2 width="51">邮　编：</td>
                <td class=S2 colspan="3" width="482"><font face="Arial"><%=rs_q("post")%></font></td>
              </tr>
              <tr bgcolor="#f6f6f6"> 
                <td class=S2 width="51">电　话：</td>
                <td class=S2 colspan="3" width="482"> 
                  <%response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3") %>
                </td>
              </tr>
              <tr> 
                <td class=S2 width="51">传　真：</td>
                <td class=S2 colspan="3" width="482"> 
                  <% response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
                </td>
              </tr>
              <tr bgcolor="#f6f6f6"> 
                <td class=S2 width="51"><font face="Arial">E-mail：</font></td>
                <td class=S2 colspan="3" width="482"><font face="Arial"><a href="mailto:<%=rs_q("email")%>"><%=rs_q("email")%></a></font></td>
              </tr>
              <tr> 
                <td class=S2 width="51">主　页：</td>
                <td class=S2 colspan="3" width="482"><font face="Arial"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a> 
                  </font></td>
              </tr>
            </table>
            <BR>
          </TD>
        </TR>
        <TR> 
          <TD><IMG height=41 src="images/MB_31.gif" width=601></TD>
        </TR>
        <TR> 
          <TD>&nbsp;</TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD>
  </TR>
  <TR> 
    <TD vAlign=top align=middle width=167 background=images/MB_38.gif height="100%"><BR>
      <BR>
    </TD>
  </TR>
  </TBODY> 
</TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <TBODY> 
  <TR> 
    <TD bgColor=#98a2c9 height=2></TD>
  </TR>
  <TR> 
    <TD> 
      <!--#include file="Copyright.htm"-->
    </TD>
  </TR>
  </TBODY> 
</TABLE>
</BODY></HTML>
<%end if%>