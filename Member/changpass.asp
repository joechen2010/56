<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->

<%
if request("action")="save" then
   if request("oldpass")<>"" then
   set rs1=server.createobject("adodb.recordset")
   sql1="select * from qyml where user='"&session("user")&"'"
   rs1.open sql1,conn,3,3
     if rs1("Pass")=trim(request("oldpass")) then
       rs1("Pass")=trim(request("newpass"))
       rs1.update
       rs1.close
	   set rs1=nothing
       response.write "<script>alert('密码修改完成！');this.location.href='changpass.asp'</script>"
     else
       response.write"<script>alert('旧密码错误，请重新输入！');history.go(-1)</script>"
     end if
   end if
end if
%>
<HTML>
<HEAD>
<TITLE>诚信物流网 会员控制中心 - 物流商人的网上家园</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
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
<META content="MSHTML 5.00.3315.2870" name=GENERATOR>
<SCRIPT language=JavaScript> 
<!--
function chk_form() 
{
var filter=/^\s*[.A-Za-z0-9_-]{5,15}\s*$/;
 if (!filter.test(document.form.oldpass.value)) 
{ alert("密码填写不正确,请重新填写！\n可使用的字符为（A-Z a-z 0-9 _ - .)\n长度5-15个字符，不要使用空格。"); 
  document.form.oldpass.focus();
  return false; 
 }
if (!filter.test(document.form.newpass.value)) 
{ alert("新密码填写不正确,请重新填写！\n可使用的字符为（A-Z a-z 0-9 _ - .)\n长度5-15个字符，不要使用空格。"); 
  document.form.newpass.focus();
  return false; 
 }
if (!filter.test(document.form.renewpass.value)) 
{ alert("确认密码填写不正确,请重新填写！\n可使用的字符为（A-Z a-z 0-9 _ - .)\n长度5-15个字符，不要使用空格。"); 
  document.form.renewpass.focus();
  return false; 
 }
if (document.form.newpass.value!=document.form.renewpass.value)
{alert("密码确认不正确，两次密码不同");
 document.form.newpass.focus();
 return false;
}
if (document.form.newpass.value==document.form.oldpass.value)
{alert("你的新密码和老密码相同，不能修改");
 document.form.newpass.focus();
 return false;
}
return true;
}
//-->
</SCRIPT>
</HEAD>
<BODY bgColor=#ffffff oncontextmenu=javascript:self.event.returnValue=false 
topMargin=0 MARGINHEIGHT="0">
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
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 修改登录密码</td>
          <td width="17%"><img src="images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="534"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> <%=session("user")%> </font> </strong>，欢迎您回来！</td>
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
                        <div align="center"><font color="#FFFFFF">用户修改密码</font></div>
                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <FORM action=changpass.asp method=post name=form 
      onsubmit="return chk_form();">
            <tr>
              <td valign="top" bgcolor="#FFFFFF">
<table width="100%"   border="0" cellpadding="0" cellspacing="0">
                    
					<tr>
                      <td valign="top">
					      <table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#B1D4F2" >
                            <tr align="left"> 
                              <td colspan="2" class="white" background="images/title_bg.gif">■ 用户修改密码</td>
                            </tr>
                            <tr> 
                              <td width="24%" align="right" bgcolor="#EDF6FF"><span class="sing"><font color="#FF0000">*</font></span> 
                                原来的密码：</td>
                              <td width="76%" align="left" bgcolor="#EDF6FF" >
                                <input class="box" maxlength=15 name=oldpass
            size=25 type=password>
                                <font color=#ff0000>*</font> </td>
                            </tr>
                            <tr> 
                              <td align="right" bgcolor="#EDF6FF"><span class="sing"><font color="#FF0000">*</font></span> 
                                新密码：</td>
                              <td align="left" bgcolor="#EDF6FF"  >
                                <input class="box" maxlength=15 name=newpass 
            size=25 type=password>
                                <font color=#ff0000>*</font> (5位以上) </td>
                            </tr>
                            <tr> 
                              <td align="right" bgcolor="#EDF6FF"> <span class="sing"><font color="#FF0000">*</font></span> 
                                确认新密码：</td>
                              <td align="left" bgcolor="#EDF6FF"  >
                                <input class="box" maxlength=15 name=renewpass
            size=25 type=password>
                                <font color=#ff0000>*</font> (5位以上) </td>
                            </tr>
                          </table>
					    
					    <table width="100%"  border="0" cellspacing="0" cellpadding="5">
                          <tr>
                            <td align="center">
							<input type="hidden" value="save" name="action">
                                <input class="btsub" name=Submit type=submit value=修改>
                                <input class="btsub" name=button onClick="javascript:location.href='index.asp';" type=button value=取消>
                              </td>
                          </tr>
                        </table>
					    <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td width="7" height="6"><img src="images/con_1.gif" width="7" height="6"></td>
                              <td background="images/con_bg1.gif"></td>
                              <td width="6" height="6"><img src="images/con_2.gif" width="6" height="6"></td>
                            </tr>
                            <tr> 
                              <td background="images/con_bg2.gif"></td>
                              <td align="left" valign="top" bgcolor="#FEFFD2"><img src="images/icon_tip.gif" width="20" height="26"> 
                                <table border=0 cellpadding=0 cellspacing=0 width="90%">
                                  <tbody> 
                                  <tr> 
                                    <td>下面是本站的建议：<br>
                                      <font 
            color=#0000ff>1.密码至少要6位，最好能在8位以上</font><br>
                                      <font 
            color=#0000ff>2.密码要定期修改，最好能在你维护你的资料时顺便修改</font><br>
                                      <font 
            color=#0000ff>3.为了有效阻止黑客软件的工具破解，请使用数字．字母．下划线的组合</font><br>
                                      <font 
            color=#0000ff>4.不使用与公司或者个人经常使用的密码，如果被攻破，则后果不堪设想．</font><br>
                                      比较好的密码如下：tj_info-987,-_freezwy8等等</td>
                                  </tr>
                                  </tbody> 
                                </table>
                              </td>
                              <td background="images/con_bg4.gif"></td>
                            </tr>
                            <tr> 
                              <td width="7" height="6"><img src="images/con_3.gif" width="7" height="6"></td>
                              <td background="images/con_bg3.gif"></td>
                              <td width="6" height="6"><img src="images/con_4.gif" width="6" height="6"></td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="10"></td>
                            </tr>
                          </table></td>
					</tr>
                  </table></td>
            </tr>
		</form>
          </table>
                        <table width="534"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="right"><a href="#top"><img src="images/icon_top.gif" alt="回到页面顶部" border="0" align="absmiddle" width="15" height="18"></a> 
                  <a href="#top">再检查一遍</a></td>
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
