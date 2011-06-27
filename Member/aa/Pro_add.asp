<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<title><%=WebName%>-会员管理平台</title>
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
alert("产品名称不能为空")
document.Form1.cpmc.focus()
document.Form1.cpmc.select() 
return
}
if (document.Form1.cpbh.value=="")
{
alert("产品编号不能为空")
document.Form1.cpbh.focus()
document.Form1.cpbh.select()
return
}
if (document.Form1.cpsb.value=="")
{
alert("产品商标不能为空")
document.Form1.cpsb.focus()
document.Form1.cpsb.select()
return
}
if (document.Form1.cpcd.value=="")
{
alert("产品产地不能为空")
document.Form1.cpcd.focus()
document.Form1.cpcd.select() 
return
}

if (document.Form1.jysm.value=="")
{
alert("简要说明不能为空！")
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
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 发布产品信息</td>
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
                        <div align="center"><font color="#FFFFFF">发布产品信息</font></div>
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
                        <div id="TipAll" > <font color="#FF0000">注：请不要发布重复产品，谢谢合作&nbsp;&nbsp;&nbsp; 
                          </font></div>
                      </td>
                            </tr>
                           </table>
                 
                  <%
dim MaxP,hydj
set rs_c=conn.execute("select count(*) from Spzs where gsid="&session("id"))
if session("hydj")=0 then
    MaxP=0
	hydj="免费会员" 
elseif session("hydj")=1 then
    MaxP=15  
	hydj="铜牌会员" 
elseif session("hydj")=2 then
    MaxP=30
	hydj="银牌会员" 
elseif session("hydj")=3 then
    MaxP=99999
	hydj="金牌会员"   
else
    MaxP=0  
end if
if rs_c(0)>=MaxP then
response.write "<p>对不起，你是"&hydj&"只能发布"&MaxP&"条信息！</p>"
end if
%>
                  <table border="0" cellspacing="1" style="border-collapse: collapse" width="100%" id="AutoNumber2" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" cellpadding="5" bgcolor="#B1D4F2">
                    <form method="POST" action="Pro_save_add.asp" name="Form1">
                      <tr> 
                        <TD align=right width=273 height="38" bgcolor="#EDF6FF">产品名称<SPAN class=word9>：</SPAN></td>
                        <TD height="38" bgcolor="#EDF6FF">
                          <input class="f11" name="cpmc" size="31">
                          &nbsp;<font color="#FF6600">*</font></td>
                      </tr>
                      <tr> 
                        <TD align="right" width="273" height="23" bgcolor="#EDF6FF">产品编号：</TD>
                        <TD width=197 height="23" bgcolor="#EDF6FF">
                          <input class="f11" name="cpbh" size="15" title=输入产品编号，如LH－002，编号不要重复，便于网上订购。>
                          &nbsp;<font color="#FF6600">*</font></td>
                      </tr>
                      <tr> 
                        <TD align="right" width="273" height="27" bgcolor="#EDF6FF">产品产地：</TD>
                        <TD width=197 height="27" bgcolor="#EDF6FF">
                          <input class="f11" name="cpcd" size="15" title=输入产品编号，如LH－002>
                        </td>
                      </tr>
                      <tr> 
                        <TD align="right" width="273" height="26" bgcolor="#EDF6FF">产品商标<span class=word9>：</span></td>
                        <TD height="26" bgcolor="#EDF6FF"> 
                          <input class="f11" name="cpsb" size="15" title=例如：桑塔纳、海尔、e点科技 >
                          &nbsp;<font color="#FF6600">*</font></TD>
                      </tr>
                      <tr> 
                        <TD align="right" width="273" height="26" bgcolor="#EDF6FF">产品图片：</td>
                        <TD height="26" bgcolor="#EDF6FF"> &nbsp;<iframe name="Picture" frameborder=0 width=430 height=20 scrolling=no src=Upload.asp></iframe> 
                          <br>
                          图片路径： 
                          <input name="picture" type="TEXT" id="picture"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100 readonly>
                        </TD>
                      </tr>
                      <tr> 
                        <TD align="right" width="273" height="52" bgcolor="#EDF6FF" valign="middle"> 
                          简要说明：<br>
                          &nbsp;(最多100字)</TD>
                        <TD align="left" height="52" bgcolor="#EDF6FF">
                          <TEXTAREA class="f11" rows="3" name="jysm" cols="51" style="font-family: 宋体" title=简要说明最多只能100个字></TEXTAREA>
                          <FONT color="#FF6600">&nbsp;*</FONT></TD>
                      </tr>
                      <tr> 
                        <TD colspan="2" align="center" height="12" bgcolor="#EDF6FF"> 
                          <input type="hidden" class="smallinput" name="jptj" value="yes">
                          <input class="f11" type="button" value="确 定" onClick="check()" name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                          <input class="f11" type="button" value="返 回" name="button">
                        </TD>
                      </tr>
                    </FORM>
                  </table>
</td>
            </tr>
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


