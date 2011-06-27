<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title><%=WebName%>-发布货运专线信息</title>
<script LANGUAGE="JavaScript">
function check()
{
if (document.Form1.province.value =="省份"||document.Form1.province2.value=="省份") 
{ 
alert("请选择区域！"); 
document.Form1.province.focus(); 
return (false); 
} 
 
if (document.Form1.content.value.length >50) 
{ 
alert('不能超过50个字!');
document.Form1.content.focus(); 
return (false); 
}


}
</SCRIPT>
<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language=JavaScript src="../inc/p_c_a2.js"></script>

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
<BODY>
<!--#include file="head.htm"-->
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
      <table width="98%"  border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 发布客运专线</td>
          <td width="17%"><img src="images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="98%"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> <%=session("user")%></font> </strong>，欢迎您回来！</td>
        </tr>
      </table>
      <table width="100%" height="497"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top">
            <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="images/title_2.jpg" bgcolor="#FFFFFF" >
                        <div align="center"><font color="#FFFFFF">发布客运专线</font></div>
                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="98%"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
			  <table width="100%"  border="0">
                             <tr>
                              <td align="left">
                        <div id="TipAll" > <font color="#FF0000">注：请不要发布重复信息，谢谢合作&nbsp;&nbsp;&nbsp;&nbsp; 
                          </font> </div>
                      </td>
                            </tr>
                           </table><%set rs_p=conn.execute("select Count(*) from info2 where GsID='"&session("user")&"'")
						   p_count=rs_p(0)
						   rs_p.close
						   set rs_p=nothing
						   if p_count>=10 then
						      response.write "<p>&nbsp;&nbsp;对不起，免费会员只能发布10条产品信息！想发更多信息请联系我们成为正式会员，谢谢！</p>&nbsp;&nbsp;咨询电话：+86-574-63809664 27861860   传真：+86-574-27860361 诚信物流网 "
						   else
						   %>
  <form method="post" action="kyzx_save.asp" name="Form1" onSubmit="javascript:return check(this);">
		                  
                    <table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="113" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                        <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">发布客运专线</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">出发地:</td>
                        <td width="81%" valign="top" height="22"> 
                          <select id="s1" name="province">
                            <option value="省份">省份</option>
                          </select>
                          <select id="s2" name="city">
                            <option>地级市</option>
                          </select>
                          <select id="s3" name="area">
                            <option>市、县级市、县</option>
                          </select>	                     </td>
                      </tr>				  
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">到达地: </td>
                        <td width="81%" valign="top" height="22">
                          <select id="t1" name="province2">
                            <option value="省份">省份</option>
                          </select>
                          <select id="t2" name="city2">
                            <option>地级市</option>
                          </select>
                          <select id="t3" name="area2">
                            <option>市、县级市、县</option>
                          </select>	                    </td>
                      </tr>
			<SCRIPT language=javascript>
            setup();

          </SCRIPT>
          <SCRIPT language=javascript>
            setup2();

         </SCRIPT>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">线路描述：</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF">
						<textarea name="content" cols="40" rows="5"></textarea>						</td>
                      </tr>
                      
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="19"> 
                          <input type="submit" value="发 布"  name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                          <input type="reset" value="重 填">						  </td>
                      </tr>
                    </table>
                  </form> 
				  </td>
            </tr>
          </table>
              <%end if%>           
            <table width="98%"  border="0" cellspacing="0" cellpadding="5">
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
</BODY>
</HTML>
