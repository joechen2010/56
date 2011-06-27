<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../Member/check.asp"-->
<html>
<head>
<title><%=WebName%>-发布车源信息</title>


<script language=JavaScript src="../inc/p_c_a.js"></script>
<script LANGUAGE="JavaScript">
function check()
{
if (document.Form1.title.value =="") 
{ 
alert("标题不能为空"); 
document.Form1.title.focus(); 
return (false); 
}
if (document.Form1.province.value =="省份") 
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
<BODY >
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="566" align="center" valign="top" bgcolor="#FFFFFF" id="main">
      <table width="98%"  border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="83%" height="38" align="left" id="pos"><a href="http://www.cx56w.com">诚信物流网</a> &gt; 发布供求信息</td>
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
  <form method="post" action="gq_save.asp" name="Form1" onSubmit="javascript:return check(this);">
		                  
                    <table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="121" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                        <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">发布库房供求信息</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">信息标题：</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600"> 
						<input name="title" type="text" maxlength="20">					  
                          </font>*</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">信息类型：</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600"> 
						<select name="infotype">
	                      <option selected value="供应">供应</option>
	                      <option value="采购">采购</option>						
						</select>						  
                          </font></td>
                      </tr>					  
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">所在地点：</td>
                        <td width="81%" valign="top" height="22"><font color="#FF6600">
                          <select id="s1" name="province">
                            <option value="省份">省份</option>
                          </select>
                          <select id="s2" name="city">
                            <option>地级市</option>
                          </select>
                          <select id="s3" name="area">
                            <option>市、县级市、县</option>
                          </select>							  
                          </font> </td>
                      </tr>
			<SCRIPT language=javascript>
            setup();
          </SCRIPT>
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
                        <td width="19%" height="22"> 内容：</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF">
						<textarea name="content" cols="40" rows="5"></textarea>						</td>
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
      </table></td>
  </tr>
</table>

</BODY>
</HTML>
