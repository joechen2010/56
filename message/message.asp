<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
Dim Gsid
Gsid=SafeRequest("Gsid",1)
if Gsid="" then
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
<TITLE><%=rs_q("qymc")%> - 物流通免费会员 - 给我留言</TITLE>
<LINK href="../images/mian.css" rel=stylesheet type=text/css> 
<STYLE type=text/css>.p14 {
	FONT-SIZE: 14px; COLOR: #000000; LINE-HEIGHT: 130%
}
.style12 {
	FONT-WEIGHT: bold; FONT-SIZE: 18px; COLOR: #ff6600
}
.style13 {
	FONT-SIZE: 14px; LINE-HEIGHT: 130%
}
.style14 {
	FONT-WEIGHT: bold; FONT-SIZE: 18px
}
</STYLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2462.0" name=GENERATOR>
</HEAD>
<BODY oncontextmenu="return false" onselectstart="return false" leftmargin="0" topmargin="0">
<!--#include file="../inc/head1.htm"-->
<table cellspacing=0 cellpadding=0 width=770 align=center border=0>
  <tbody> 
  <tr> 
    <td class=S height=30>您当前位置：<a href="../">首页</a> &gt;&gt; 留言询价</td>
  </tr>
  </tbody> 
</table>
<%if session("user")="" then %>
      <table width="755" border="0" align="center" cellpadding="4" cellspacing="0">
				<tr>
					<td height="35" valign="bottom">
						<marquee onMouseOver="this.stop()" onMouseOut="this.start()" scrollamount="5" direction="left"
							behavior="alternate" width="620">
							<font color="#003333"><span class="P14"><strong>你目前进行的操作，需要“登录”后才能继续……</strong></span>
							</font>
						</marquee>
					</td>
				</tr>
			</table>
			<table width="755" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr>
					<td width="384" height="147"><img src="../Login/images/login02.gif" width="384" height="147"></td>
					
      <td width="371" background="../Login/images/login03.gif">&nbsp;</td>
				</tr>
				<tr>
					<td height="159" valign="bottom" background="../Login/images/login05.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="22%"></td>
								
          <td width="78%"><a href="../Reg/User_Reg.asp"><img src="../Login/images/login09.gif" width="219" height="61" border="0"></a></td>
							</tr>
							<tr>
								<td height="60" colspan="2"></td>
							</tr>
						</table>
					</td>
					<td valign="bottom" background="../Login/images/login06.gif">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="13%" height="80"><FONT face="宋体"></FONT></td>
            <td width="87%">
              
            <table width="100%" border="0" cellspacing="3" cellpadding="3">
              <form name=form action=../Login/check2.asp method=post>
                <tr> 
                  <td width="16%" height="35" align="center">用户名</td>
                  <td width="39%"> 
                    <input name="UserID" type="text" id="Accounts" class="formtext" style="width:100px;" />
                  </td>
                  <td width="45%" rowspan="2" align="left"> 
                    <input type="image" name="btn_Ok" id="btn_Ok" src="../Login/images/login10.gif" alt="" border="0" style="height:52px;width:67px;" />
                  </td>
                </tr>
                <tr> 
                  <td height="35" align="center">密 码</td>
                  <td> 
                    <input name="Pwd" type="password" id="Pwd" class="formtext" style="width:100px;" />
                  </td>
                </tr></form>
              </table>
            </td>
          </tr>
          <tr> 
            <td width="13%"></td>
            <td height="30" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;<a href="../Login/getAnswer.asp"><font color="#000000"><u>忘记密码怎么办？</u></font></a> 
            </td>
          </tr>
        </table>
					</td>
				</tr>
				<tr>
					<td height="40" style="HEIGHT: 40px"><img src="../Login/images/login07.gif" width="384" height="41"></td>
					<td style="HEIGHT: 40px"><img src="../Login/images/login08.gif" width="371" height="41"></td>
				</tr>
			</table>
	
<%else %>
<table cellspacing=0 cellpadding=0 width=770 align=center 
background=images/inquiry_t_bg.gif border=0>
  <tbody> 
  <tr> 
    <td><img height=60 src="images/inquiry_t_l.gif" border=0></td>
    <td width=13><img height=60 src="images/inquiry_t_r.gif" 
    width=13></td>
  </tr>
  </tbody>
</table>
<table cellspacing=0 cellpadding=0 width=770 align=center border=0>
  <tbody> 
  <tr> 
    <td width=5 background=images/pro_r_l_bg.gif><img height=1 
      src="images/spacer.gif" width=5></td>
    <td align=middle>        
      <table class=style cellspacing=5 cellpadding=0 width="100%" 
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

if (document.Form1.F_Fax.value =="")
{
alert("请填写传真号码！");
document.Form1.F_Fax.focus();
document.Form1.F_Fax.select();
return false;
document.Form1.submit()
}
</script>
                        <form method="POST" action="messagesave.asp" name="Form1"  onSubmit="return CheckForm();">
                          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber10">
                            <tr> 
                              
                  <td width="100%" height="267"> 
                    <table cellspacing=4 cellpadding=1 width="90%" align=center border=0>
                      <tbody> 
                      <tr> 
                        <td class=S 
            height=25>您的信息将被及时、准确地发送给对方，我们承诺未经您的允许不会以任何方式将您的个人信息透露给第三方。</td>
                      </tr>
                      </tbody> 
                    </table>
                    <table cellspacing=1 cellpadding=6 width="90%" align=center 
      bgcolor=#ff8200 border=0>
                      <tr bgcolor="#ffe8d0" align="left"> 
                        <td colspan="2" class=yin1>填写询价信息：</td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right">收件人公司：</td>
                        <td width="503"><b class="L"><%=rs_q("qymc")%></b></td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right"><font color="#000000">收件人：</font></td>
                        <td width="503"><%=rs_q("name")%></td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right"><font color="#FF0000">* </font><font color="#000000"></font><font color="#000000">主题：</font></td>
                        <td width="503"> 
                          <input type="text" name="subject" value="询问业务，请速联系">
                        </td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right"><font color="#FF0000">* </font>内容：</td>
                        <td width="503"> 
                          <textarea cols=50 name=content rows=6 style="border-style: solid; border-width: 1"></textarea>
                        </td>
                      </tr>
                      <tbody> 
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right"><font color=#ff6600>*</font> 
                          您的公司：</td>
                        <td width="503"> 
                          <input type="text" size="33" name="F_Company" value=<%=session("qymc")%>>
                        </td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right"><font color=#ff6600>*</font> 
                          您的名字：</td>
                        <td width="503"> 
                          <input type="text" size="33" name="F_Name" value=<%=session("name")%>>
                        </td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right"><font color=#ff6600>*</font> 
                          您的邮件：</td>
                        <td width="503"> 
                          <input type="text" size="33" name="F_Email" value=<%=session("email")%>>
                        </td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right"><font color=#ff6600>*</font> 
                          联系地址：</td>
                        <td width="503"> 
                          <input type="text" size="33" name="F_address" value=<%=session("address")%>>
                        </td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right"><font color=#ff6600>*</font> 
                          邮&nbsp;&nbsp;&nbsp;&nbsp;编:</td>
                        <td width="503"> 
                          <input type="text" size="33" name="F_Post" value=<%=session("post")%>>
                        </td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right"> <font color=#ff6600>*</font> 
                          电&nbsp;&nbsp;&nbsp;&nbsp;话：</td>
                        <td width="503"> 
                          <input type="text" size="33" name="F_Tel" value=<%=session("phone")%>>
                        </td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td width="96" align="right">&nbsp; <font color=#ff6600>*</font> 
                          传&nbsp;&nbsp;&nbsp;&nbsp;真:</td>
                        <td width="503"> 
                          <input type="text" size="33" name="F_Fax" value=<%=session("fax")%>>
                        </td>
                      </tr>
                      <tr bgcolor="#f9f9f9"> 
                        <td colspan="2"> 
                          <input type="hidden" name="F_User" value="<%=session("user")%>">
                          <input type="hidden" name="T_User" value="<%=rs_q("user")%>">
                          <input type="hidden" name="T_Name" value="<%=rs_q("name")%>">
                          <input type="hidden" name="T_Company" value="<%=rs_q("qymc")%>">
                          <input type="hidden" name="T_Email" value="<%=rs_q("email")%>">
                        </td>
                      </tr>
                      <tr align="CENTER" bgcolor="#f9f9f9"> 
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
    <td width=6 background=images/pro_r_r_bg.gif><img height=8 
      src="images/spacer.gif" width=6></td>
  </tr>
  </tbody>
</table>
<table cellspacing=0 cellpadding=0 width=770 align=center 
background=images/pro_r_b_bg.gif border=0>
  <tbody> 
  <tr> 
    <td align=middle width=11><img height=17 
      src="images/pro_r_b_l.gif" width=11></td>
    <td align=middle><img height=1 src="images/spacer.gif" 
width=1></td>
    <td align=middle width=13><img height=17 
      src="images/pro_r_b_r.gif" width=13></td>
  </tr>
  </tbody>
</table>
<%end if%>
<!--#include file="../inc/down.htm"-->      
</BODY></HTML>
<%end if%>