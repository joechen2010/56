<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../Member/check.asp"-->
<html>
<head>
<title><%=WebName%>-发布货源信息</title>

<script LANGUAGE="JavaScript">
function check()
{
if (document.Form1.province.value =="省份"||document.Form1.province2.value=="省份") 
{ 
alert("请选择区域！"); 
document.Form1.province.focus(); 
return (false); 
} 
 
if (document.Form1.carload.value =="") 
{ 
alert("货物重量不能为空"); 
document.Form1.carload.focus(); 
return (false); 
}
if (document.Form1.carlong.value =="") 
{ 
alert("车长不能为空"); 
document.Form1.carlong.focus(); 
return (false); 
}
if (document.Form1.content.value.length >50) 
{ 
alert('不能超过50个字!');
document.Form1.content.focus(); 
return (false); 
}

}
</script>
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
<link href="../Member/images/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
</head>
<BODY>
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="566" align="center" valign="top" bgcolor="#FFFFFF" id="main">
      <table width="98%"  border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="83%" height="38" align="left" id="pos"><a href="http://www.cx56w.com">诚信物流网</a> &gt; 发布货源信息</td>
          <td width="17%">&nbsp;</td>
      </tr>
    </table>
	  <table width="100%" height="497"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top">
            <table width="98%"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
			  <table width="100%"  border="0">
                             <tr>
                              <td align="left">
                        <div id="TipAll" > <font color="#FF0000">注：请不要发布重复信息，谢谢合作&nbsp;&nbsp;&nbsp;&nbsp; 
                          </font> </div>                      </td>
                            </tr>
                           </table><%set rs_p=conn.execute("select Count(*) from info2 where GsID='"&session("user")&"'")
						   p_count=rs_p(0)
						   rs_p.close
						   set rs_p=nothing
						   if p_count>=10 then
						      response.write "<p>&nbsp;&nbsp;对不起，免费会员只能发布10条产品信息！想发更多信息请联系我们成为正式会员，谢谢！</p>&nbsp;&nbsp;咨询电话：+86-574-63809664 27861860   传真：+86-574-27860361 诚信物流网 "
						   else
						   %>
  <form method="post" action="goods_save.asp" name="Form1" onSubmit="javascript:return check(this);">
		                  
                    <table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="121" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                        <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">发布货源信息</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">出发地:</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600"> 
                          <select id="s1" name="province">
                            <option value="省份">省份</option>
                          </select>
                          <select id="s2" name="city">
                            <option>地级市</option>
                          </select>
                          <select id="s3" name="area">
                            <option>市、县级市、县</option>
                          </select>
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">到达地: </td>
                        <td width="81%" valign="top" height="22"><font color="#FF6600">
                          <select id="t1" name="province2">
                            <option value="省份">省份</option>
                          </select>
                          <select id="t2" name="city2">
                            <option>地级市</option>
                          </select>
                          <select id="t3" name="area2">
                            <option>市、县级市、县</option>
                          </select>
                          </font> </td>
                      </tr>
			<SCRIPT language=javascript>
            setup();

          </SCRIPT>
          <SCRIPT language=javascript>
            setup2();

         </SCRIPT>					  
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">货物种类:</td>
                        <td width="81%" valign="top" height="22"> 
                        <select name="cartype" id="cartype">
                          		<option value="有货">有货</option>
		<option value="整车货">整车货</option>
		<option value="零担货">零担货</option>
		<option value="重货">重货</option>
		<option value="泡货">泡货</option>
		<option value="漂货">漂货</option>
		<option value="普货">普货</option>
		<option value="化工产品">化工产品</option>
		<option value="设备">设备</option>
		<option value="蔬菜">蔬菜</option>
		<option value="建材">建材</option>
		<option value="木材">木材</option>
		<option value="食品">食品</option>
		<option value="粮食">粮食</option>
		<option value="冷冻食品">冷冻食品</option>
		<option value="农副产品">农副产品</option>
		<option value="水果">水果</option>
		<option value="矿石">矿石</option>
		<option value="煤">煤</option>
		<option value="医药">医药</option>
		<option value="家具">家具</option>
		<option value="纸箱">纸箱</option>
		<option value="危化品">危化品</option>
		<option value="随便装">随便装</option>
		<option value="配件">配件</option>
		<option value="小货">小货</option>
		<option value="捎带货">捎带货</option>
	                </select>						</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">货物重量:</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600">
                          <input name=carload id=carload maxlength="4">
                          &nbsp;
                          </font>吨</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">车长:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                          <input name=carLong id=carlong maxlength="4">
                          &nbsp;
                          </font>米</td>
                      </tr>
                      <tr bgcolor="#EDF6FF">
                        <td width="19%" height="22">有效期:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                        <select name="validate" id="validate">
						
		                   <option value="5">5天</option>
		                   <option value="15">15天</option>
		                   <option value="30">30天</option>
						   <option value="60">60天</option>
						   <option value="90">90天</option>
						   <option value="120">120天</option>
						   </select>
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">说明：</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF">
						<textarea name="content" cols="40" rows="5"></textarea>						</td>
                      </tr>
                      
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="19"> 
                          <input type="submit" value="发 布"  name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                        <input type="button" onClick="javascript:history.back(1)" value="返回">
					   </td>
                      </tr>
                    </table>
                  </form> <%end if%>               </td>
            </tr>
          </table>
                        
            </td>
        </tr>
      </table></td>
  </tr>
</table>
</BODY>
</HTML>
