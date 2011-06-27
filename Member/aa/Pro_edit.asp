<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<%
if not isEmpty(request("id")) then
    id=request("id")
else
    response.write "参数错误！"
    response.end
end if
sql="select * from spzs where id="&cstr(id) 
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<html>
<head>
<title>修改产品信息</title>
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
<SCRIPT language=javascript src="images/zyok.JS"></SCRIPT>
<script language="javascript">
function show_sader(mylink)
{
window.open(mylink,'','top=140,left=200,width=480,height=200,scrollbars=yes')
}
</script>
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
</head>
<BODY text=#000000 leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" bgcolor="#F7FBFF">
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
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 修改产品信息</td>
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
          <td align="center" valign="top"> 
            <table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td> 
                  <table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="images/title_2.jpg" bgcolor="#FFFFFF" > 
                        <div align="center"><font color="#FFFFFF">修改产品信息</font></div>
                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr> 
                <td valign="top" bgcolor="#FFFFFF"> 
                  <table width="100%"  border="0">
                    <tr> 
                      <td align="left"> 
                        <div id="TipAll" > 请输入所有标有*的项目。 </div>
                      </td>
                    </tr>
                  </table>
                  <table border="0" cellspacing="1" style="border-collapse: collapse" width="100%" id="AutoNumber2" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" cellpadding="5" bgcolor="#B1D4F2">
                    <FORM method="post" Action="Pro_Edit_save.asp" Name="Form1">
                      <input type="hidden" name="id" value="<%=id%>">
                      <input type="hidden" name="gsid" value="<%=rs("gsid")%>">
                      <tr bgcolor="#EDF6FF"> 
                        <TD align=right width=205 height="27"><font color="#FF6600">*</font> 
                          产品名称<SPAN class=word9>：</SPAN></td>
                        <TD height="27"> 
                          <input class="f11" name="cpmc" size="30" value="<%=rs("cpmc")%>">
                          &nbsp;</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <TD align="right" width="205" height="23"><font color="#FF6600">*</font> 
                          产品编号：</TD>
                        <TD width=265 height="23" bgcolor="#EDF6FF"> 
                          <input class="f11" name="cpbh" size="30" title=输入产品编号，如LH－002 value="<%=rs("cpbh")%>">
                          &nbsp;</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <TD align="right" width="205" height="27"><font color="#FF6600">*</font> 
                          产品商标<span class=word9>：</span></TD>
                        <TD width=265 height="27"> 
                          <input class="f11" name="cpsb" size="30" title=例如：桑塔纳、海尔、e点科技 value="<%=rs("cpsb")%>" >
                          &nbsp;</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <TD align="right" width="205" height="27"><font color="#FF6600">*</font> 
                          产品产地：</TD>
                        <TD width=265 height="27"> 
                          <input class="f11" name="cpcd" size="30" title=输入产品编号，如LH－002 value="<%=rs("cpcd")%>">
                        </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <TD align="right" width="205" height="26"><font color="#FF6600">*</font> 
                          产品图片：</td>
                        <TD height="26" bgcolor="#EDF6FF"> &nbsp;<iframe name="Picture" frameborder=0 width=400 height=20 scrolling=no src=Upload.asp></iframe> 
                          <br>
                          图片路径： 
                          <input name="picture" type="TEXT" id="picture"  style="background-color:ffffff;color:000000;border: 1 double" value="<%=rs("Picture")%>" size=34 maxlength=100 readonly>
                        </TD>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <TD align="right" width="205" height="52" valign="middle"> 
                          <font color="#FF6600">*</font> 简要说明：<br>
                          &nbsp;(最多100字)</TD>
                        <TD align="left" height="52"> 
                          <TEXTAREA class="f11" rows="6" name="jysm" cols="50" style="font-family: 宋体" title=简要说明最多只能100个字><%=dvHTMLCode(rs("jysm"))%></TEXTAREA>
                          <FONT color="#FF6600">&nbsp;</FONT></TD>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <TD colspan="2" align="center" height="12"> 
                          <input class="f11" type="button" value="确 定" onClick="check()" name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                          <input class="f11" type="button" value="返 回" name="button" onClick="history.go(-1)">
                          <input type="hidden" class="smallinput" name="jptj" value="yes" checked>
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
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!--#include file="bottom.htm"-->
</body>
</html>