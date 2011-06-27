<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<title><%=WebName%>-发布车源信息</title>

<SCRIPT language=javascript  src="JSROOT/CITY.JS"></SCRIPT>
<SCRIPT language=javascript  src="JSROOT/CITY1.JS"></SCRIPT>
<script LANGUAGE="JavaScript">
function check()
{
if (document.Form1.carnum.value =="") 
{ 
alert("车牌号不能为空！"); 
document.Form1.carnum.focus(); 
return (false); 
} 
if (document.Form1.carpp.value =="") 
{ 
alert("品牌不能为空！"); 
document.Form1.carpp.focus(); 
return (false); 
}

if (document.Form1.cartype.value =="") 
{ 
alert("载重不能为空"); 
document.Form1.cartype.focus(); 
return (false); 
} 
if (document.Form1.carload.value =="") 
{ 
alert("载重不能为空"); 
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
</SCRIPT>

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
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 发布车源信息</td>
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
                        <div align="center"><font color="#FFFFFF">发布车源信息</font></div>
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
  <form method="post" action="files_save.asp" name="Form1" onSubmit="javascript:return check(this);">
		                  
                    <table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="121" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                        <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">发布车源信息</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">车牌号:</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600"> 
						 <input name=carnum type="text" maxlength="10">
                        </font>*</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">品牌: </td>
                        <td width="81%" valign="top" height="22"><font color="#FF6600">
						 <input name=carpp type="text" maxlength="10">						  
                          </font> * </td>
                      </tr>					  
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">车型:</td>
                        <td width="81%" valign="top" height="22"> 
                        <select name=cartype id="cartype">
		                   <option value="有车">有车</option>
		                   <option value="普通车">普通车</option>
		                   <option value="软蓬车">软蓬车</option>
		                   <option value="半挂车">半挂车</option>
		                   <option value="厢式车">厢式车</option>
		                   <option value="冷藏车">冷藏车</option>
		                   <option value="集装箱">集装箱</option>
		                   <option value="平板车">平板车</option>
		                   <option value="起重车">起重车</option>
		                   <option value="自卸车">自卸车</option>
		                   <option value="油罐车">油罐车</option>
		                   <option value="后八轮">后八轮</option>
		                   <option value="前四后八">前四后八</option>
		                   <option value="双桥车">双桥车</option>
		                   <option value="加长挂">加长挂</option>
		                   <option value="半封车">半封车</option>
		                   <option value="高栏车">高栏车</option>
		                   <option value="保温车">保温车</option>
	                </select>
						</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">载重:</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600">
                          <input name=carload type="text" maxlength="4">
                        </font>*</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">车长:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                          <input name=carlong type="text" maxlength="4">
                        </font>*</td>
                      </tr>
                      <tr bgcolor="#EDF6FF">
                        <td width="19%" height="22">出厂日期:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
						<input name=time type="text" maxlength="10">
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF">
                        <td width="19%" height="22">主要司机:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
						<input type="text" name=siji value="<%=session("siji")%>" readonly>
                          </font></td>
                      </tr>					  
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">该车状况：</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF">
						<textarea name="content" cols="40" rows="5"></textarea>
						</td>
                      </tr>
                      
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="19"> 
                          <input type="submit" value="发 布"  name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                          <input type="reset" value="重 填">                        </td>
                      </tr>
                    </table>
                  </form> <%end if%>               </td>
            </tr>
          </table>
                        
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
