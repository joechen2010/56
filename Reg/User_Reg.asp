<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/config.asp"-->
<%
dim RegUserName,txtVerify
'RegUserName=request.form("userid")
RegUserName="**********************"
session("RegUserName")=RegUserName

'if RegUserName="" or strLength(RegUserName)>20 or strLength(RegUserName)<4 then
	'founderr=true
	'errmsg=errmsg & "请输入用户名(不能大于20小于5)\n"
'else
  	'if Instr(RegUserName,"<")>0 or Instr(RegUserName,">")>0 or Instr(RegUserName,"=")>0 or Instr(RegUserName,"%")>0 or Instr(RegUserName,chr(32))>0 or Instr(RegUserName,"?")>0 or Instr(RegUserName,"&")>0 or Instr(RegUserName,";")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,"'")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,chr(34))>0 or Instr(RegUserName,chr(9))>0 or Instr(RegUserName,"")>0 or Instr(RegUserName,"$")>0 then
		'errmsg=errmsg+"用户名中含有非法字符\n"
		'founderr=true
	'end if
'end if

if founderr=false then
	dim sqlReg,rsReg,sqladminname,rsadminname

	sqladminname="select * from Manage_user where UserName='" & RegUserName & "'"
	set rsadminname=conn.execute(sqladminname)
	
	if not(rsadminname.bof and rsadminname.eof) then
		founderr=true
		errmsg=errmsg & "朋友，想冒充管理员？换一个用户名吧老大！\n"
		session("RegUserName")=""
	'else
	
    'set rs=server.CreateObject ("adodb.recordset")
      'sql="select * from qyml where user='"&RegUserName&"'"
      'rs.open sql,conn,1,3
      'if not (rs.eof and rs.bof) then
           'founderr=true
		  ' errmsg=errmsg & "你注册的用户已经存在！请换一个用户名再试试！\n"
		 '  session("RegUserName")=""
    '  end if
	 ' rs.close
	  'set rs=nothing
	end if
end if

if founderr=false then
	call RegNext()
else
	call WriteRegErrmsg()
end if 

Sub RegNext()
%>
<html>
<head>
<title>会员注册——诚信物流网</title>
<%
dim rs
dim sql
set rs=server.createobject("adodb.recordset")
%>


<script Language="JavaScript"> 
function FormCheck()
{ 
if (document.UserReg.pass.value =="") 
{
alert("请填写您的密码！");
document.UserReg.pass.focus();
return false; 
}
if(document.UserReg.confirmPassword.value==""){
alert("请输入您的确认密码！");
document.UserReg.confirmPassword.focus();
return false;
}
var filter=/^\s*[.A-Za-z0-9_-]{5,15}\s*$/;
if (!filter.test(document.UserReg.pass.value)) { 
alert("密码填写不正确,请重新填写！可使用的字符为（A-Z a-z 0-9 _ - .)长度不小于5个字符，不超过15个字符，注意不要使用空格。"); 
document.UserReg.pass.focus();
document.UserReg.pass.select();
return false; 
} 
if (document.UserReg.pass.value!=document.UserReg.confirmPassword.value ){
alert("两次填写的密码不一致，请重新填写！"); 
document.UserReg.pass.focus();
document.UserReg.pass.select();
return false; 
} 
if (document.UserReg.name.value =="") 
{ 
alert("请输入您的姓名！"); 
document.UserReg.name.focus(); 
return (false); 
} 

if (document.UserReg.qymc.value =="") 
{ 
alert("请输入公司名称！"); 
document.UserReg.qymc.focus(); 
return (false); 
} 

if (document.UserReg.siji.value =="") 
{ 
alert("请输入司机名称！"); 
document.UserReg.siji.focus(); 
return (false); 
} 

if (document.UserReg.Ys_m_M_province.value =="省份") 
{ 
alert("请选择区域！"); 
document.UserReg.Ys_m_M_province.focus(); 
return (false); 
} 
if (document.UserReg.Ys_m_M_city.value =="") 
{ 
alert("请选择城镇名！"); 
document.UserReg.Ys_m_M_city.focus(); 
return (false); 
} 
if (document.UserReg.Ys_m_M_area.value =="") 
{ 
alert("请选择城镇名！"); 
document.UserReg.Ys_m_M_area.focus(); 
return (false); 
}
if (document.UserReg.Ys_m_M_address.value =="") 
{ 
alert("请输入您的联系地址！"); 
document.UserReg.Ys_m_M_address.focus(); 
return (false); 
} 

if (document.UserReg.Ys_m_M_phone2.value ==""||document.UserReg.Ys_m_M_phone3.value =="") 
{ 
alert("请输入您的联系电话！"); 
document.UserReg.Ys_m_M_phone2.focus(); 
return (false); 
} 
if (document.UserReg.mobile.value =="") 
{ 
alert("请输入您的手机号码！"); 
document.UserReg.mobile.focus(); 
return (false); 
} 
mobile
var regexp = /^[0123456789]{1,6}$/g; 
if (regexp.test(document.UserReg.post.value)!=1)
{
alert("请输入正确的邮政编码！");
document.UserReg.post.focus(); 
document.UserReg.post.select();
return (false);
}
//var filter2=/^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
//if (!filter2.test(document.UserReg.email.value)) { 
//alert("邮件地址不正确,请重新填写！"); 
//document.UserReg.email.focus();
//document.UserReg.email.select();
//return false; 
//} 
if (document.UserReg.qyjj.value =="")
{ 
alert("请输入公司简介！"); 
document.UserReg.qyjj.focus(); 
return (false); 
} 
if( document.UserReg.txtVerify.value =="") {
		alert("\请输入您的认证码 !");
		document.UserReg.txtVerify.focus();
		return false;
}
if( document.UserReg.txtVerify.value.length!=4 || fIsNumber(document.UserReg.txtVerify.value,"1234567890")!=1) {
		alert("\验证码只能为右边图片所示4位数字!");
		document.UserReg.txtVerify.focus();
		return false;
	}
} 

//****判断是否是Number.
function fIsNumber (sV,sR) {
	var sTmp;
	if(sV.length==0){ return (false);}
	for (var i=0; i < sV.length; i++){
		sTmp= sV.substring (i, i+1);
		if (sR.indexOf (sTmp, 0)==-1) {return (false);}
	}
	return (true);
}
//***
</script> 
<script language=JavaScript src="../inc/p_c_a.js"></script>
<style type="text/css">
<!--
.STYLE19 {color: #000000}
.STYLE20 {color: #FF0000}
-->
</style>
<head>
<LINK href="images/css.css" type="text/css" rel="stylesheet">
<style type="text/css">
.p14 { FONT-SIZE: 14px; COLOR: #000000; LINE-HEIGHT: 130% }
	.S { FONT-SIZE: 12px }
	.style16 { FONT-WEIGHT: bold }
	.button { BORDER-RIGHT: #ffffff thin; BACKGROUND-POSITION: center center; BORDER-TOP: #ffffff thin; FONT-SIZE: 12px; BORDER-LEFT: #ffffff thin; WIDTH: 50px; COLOR: #ffffff; BORDER-BOTTOM: #ffffff thin; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 20px; BACKGROUND-COLOR: #999999 }
	.button01 { BORDER-RIGHT: #9c9a9c 1px solid; BACKGROUND-POSITION: center center; BORDER-TOP: #9c9a9c 1px solid; FONT-SIZE: 12px; MARGIN: 1px; BORDER-LEFT: #9c9a9c 1px solid; WIDTH: 100px; COLOR: #666666; BORDER-BOTTOM: #9c9a9c 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: auto; BACKGROUND-COLOR: #cccccc }
	.button02 { BORDER-RIGHT: #9c9a9c 1px solid; BACKGROUND-POSITION: center center; BORDER-TOP: #9c9a9c 1px solid; FONT-SIZE: 12px; MARGIN: 1px; BORDER-LEFT: #9c9a9c 1px solid; WIDTH: 50px; COLOR: #666666; BORDER-BOTTOM: #9c9a9c 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 20px; BACKGROUND-COLOR: #cccccc }
	.button04 { BORDER-RIGHT: #aaaaaa 1px solid; BORDER-TOP: #fefefe 1px solid; FONT-SIZE: 12px; BACKGROUND: #ececec; BORDER-LEFT: #fefefe 1px solid; WIDTH: 380px; COLOR: #000000; PADDING-TOP: 1px; BORDER-BOTTOM: #aaaaaa 1px solid; WHITE-SPACE: normal; HEIGHT: 23px }
	.button03 { BORDER-RIGHT: #aaaaaa 1px solid; BORDER-TOP: #fefefe 1px solid; FONT-SIZE: 12px; BACKGROUND: #ececec; BORDER-LEFT: #fefefe 1px solid; COLOR: #000000; PADDING-TOP: 1px; BORDER-BOTTOM: #aaaaaa 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 23px }
	.style17 { FONT-WEIGHT: bold; COLOR: #cc3333 }
	</style>
</head>

<BODY onLoad="setup()">
<table cellspacing=0 cellpadding=0 width=776 align=center border=0>
  <tbody> 
  <tr> 
    <td valign=bottom> 
      <table cellspacing=0 cellpadding=0 width=776 border=0>
        <tbody> 
        <tr> 
          <td> 
            <table style="BORDER-BOTTOM: #5286b5 1px solid" height=70 cellspacing=0 
      cellpadding=0 width=776 border=0>
              <tbody> 
              <tr> 
                <td width=265 height=64><a href="../"><img height=65 
            alt=返回首页 src="images/logo.gif" width=265 align=bottom 
            border=0></a></td>
                <td valign=bottom align=middle width=511> 
                  <table height=50 cellspacing=0 cellpadding=0 width=489 border=0>
                    <tbody> 
                    <tr> 
                      <td width=180><img height=50 src="images/title1_re1.gif" 
                  width=180></td>
                      <td width=180><img height=50 src="images/title1_re2.gif" 
                  width=180></td>
                      <td width=129><img height=50 src="images/title_re3.gif" 
                  width=129></td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
            <table cellspacing=0 cellpadding=10 width="100%" border=0>
              <tbody> 
              <tr> 
                <td align=middle width="12%"><img height=30 
                  src="images/htxt_01.gif" 
        width=500></td>
                <td width="88%"><span class=S><span 
            class=style17>　<font 
      color=#ff0000><strong>24小时客服热线：(0)13116601816</strong></font></span></span></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<table cellspacing=0 width=768 border=0 align="center">
  <tbody> 
  <tr> 
    <td colspan=2><img height=80 
      alt=只要注册了本站的会员，您就会得到属于您自己的网站，可以在自己的网站上发布空车信息、货源信息及车辆档案，让货主、车主主动找到你。 
      src="images/56hy.gif" width=776></td>
  </tr>
  </tbody> 
</table>
<table width="776" border="0" align="center" cellspacing="0" cellpadding="0" class=huang12>
  <tr> 
    <td width="231" valign="top" class=huang12 style="LINE-HEIGHT: 150%" bgcolor="#FFFAE1"> 
      <p align=center><b><font size=4>如何填写会员资料 </font></b></p>
      <ol>
        <li class="STYLE19">“ * ”表示必填项． 
        <li class="STYLE19">所在地区必须精确到区县级(第三个下拉框必选)． 
        <li class="STYLE19">推荐会员编号是别的会员推荐您加入的那个会员的编号,如果没有人推荐,可以不填. 
        <li class="STYLE19">您填写的单位介绍将出现在您的网页上，最好填写完整，如果不填，您的网页上的单位介绍一栏将是空白，这会影响您的企业形象． 
        <li class="STYLE19">如果您有电子邮件信箱地址，最好填上，如果您忘了登录密码，系统会将密码发到您的邮箱中,也以便我们或其他人与您及时沟通． 
        <li class="STYLE19">注册成功后，系统将为您分配一个会员编号，这是会员的唯一标识，请牢记． 
        <li class="STYLE19">注册成功后，您就可以登录网站了，如果您选择了＂今后自动登录＂，则以后再访问我们的网站时，将自动登录，不必每次都输入会员编号和密码． 
        <font color="#FF0000"><li>注意：您注册成功后可能在网站首页不能立即看到您的单位名称或负责人，这并不表明您注册不成功，因为首页有一个延迟（最多5分钟），请不要再次重复注册！
        </li></font> 
      </ol>
    </td>
    <td width="545"> 
      <table border="0" bordercolor="#111111" cellspacing="0" cellpadding="0" align="center" width="99%" valign="top">
        <tr> 
    <td> 
      <form method="POST" action="User_RegOK.asp" name="UserReg" onSubmit="return FormCheck();" >
        <input type="hidden" name="userstate" value="new">
        <input type="hidden" name="_fom_grp_" value="Ys"/>
        <input type="hidden" name="M" value="_m_M_"/>
              <table width="96%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#BCBCBC">
                <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>用 户 名:</span></div></td>
                  <td width="33%" bgcolor="#FFFFFF"><span class=p141 STYLE18> 
                    <input class=f11 maxlength=20 name="Ys_m_M_user" size="20" value="">
                    </span> <span class="STYLE20">*</span></td>
            <td width="15%" bgcolor="#F7F7F7"><div align="right"><span 
            class=style16>您的姓名:</span></div></td>
            <td width="35%" bgcolor="#FFFFFF"><input class=f11 maxlength=100 name=name size="20">
              <font color=#990000> </font><span class=red>*</span></td>
          </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>密　　码:</span></div></td>
                  <td width="33%" bgcolor="#FFFFFF"> 
                    <input maxlength=20 name=pass type=password size="20">
                    <span class=red STYLE18>*</span></td>
            <td width="15%" bgcolor="#F7F7F7"><div align="right"><span 
            class=style16>确认密码:</span></div></td>
            <td width="35%" bgcolor="#FFFFFF"><input class=f11 maxlength=20 name=confirmPassword type=password size="20">
              <span class=p141> </span><span class=red>*</span></td>
          </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>性　　别:</span></div></td>
            <td colspan="3" bgcolor="#FFFFFF"><input class=f11 name=ch type=radio value="先生" checked>
先生&nbsp;&nbsp;
<input class=f11 name=ch type=radio value="女士">
女士<span class=red></span></td>
            </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>公司名称:</span></div></td>
            <td colspan="3" bgcolor="#FFFFFF"><span class=p141><font face=宋体><font face=宋体>
              <input  
                  name=qymc>
              <span class=red>*</span></font></font></span></td>
            </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>产业种类:</span></div></td>
                  <td width="33%" bgcolor="#FFFFFF">
                    <%
        sql = "select * from Class_1"
        rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	response.write "请先添加栏目。"
	response.end
	else
%>
                    <select name="SortID">
                <option selected value="<%=trim(rs("sort"))%>"><%=trim(rs("sort"))%></option>
                <%      dim selclass
         selclass=rs("SortID")
        rs.movenext
        do while not rs.eof
%>
                <option value="<%=trim(rs("sort"))%>"><%=trim(rs("sort"))%></option>
                <%
        rs.movenext
        loop
	end if
        rs.close
%>
              </select>
                    <font color=#FF0000>*</font></td>
            <td width="15%" bgcolor="#F7F7F7"><div align="right"><span class="style16">主要司机:</span></div></td>
            <td width="35%" bgcolor="#FFFFFF"><input type="text" name="siji"></td>
          </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>公司所在地:</span></div></td>
            <td colspan="3" bgcolor="#FFFFFF"><select id="s1" name="Ys_m_M_province">
              <option value="省份">省份</option>
            </select>
              <select id="s2" name="Ys_m_M_city">
                <option>地级市</option>
              </select>
              <select id="s3" name="Ys_m_M_area">
                <option>市、县级市、县</option>
              </select>
              <span class=red>*</span></td>
            </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>公司地址:</span></div></td>
            <td colspan="3" bgcolor="#FFFFFF"><input id=address  
                  name="Ys_m_M_address">
&nbsp; <span class=red>*</span> 只需填写街道、门牌号</td>
            </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>固定电话:</span></div></td>
            <td colspan="3" bgcolor="#FFFFFF"><span class=p141>
              <input 
                  id=Ys_m_M_phone1 style="with: 37px" value=+86 name="Ys_m_M_phone1" type="hidden">
            </span><span class=p141>
            <input 
                  id=Ys_m_M_phone2 style="with: 53px" name="Ys_m_M_phone2" size="4">
            </span>-<span class=p141>
            <input 
                  id=Ys_m_M_phone3 style="with: 103px" name="Ys_m_M_phone3">
            </span><span class=red>* </span>区号-电话号码</td>
            </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>传　　真:</span></div></td>
            <td colspan="3" bgcolor="#FFFFFF"><span class=p141>
              <input 
                  id=tel1 style="with: 37px" value=+86 name="Ys_m_M_fax1" type="hidden">
            </span><span class=p141>
            <input 
                  id=tel2 style="with: 53px" name="Ys_m_M_fax2" size="4">
            </span>-<span class=p141>
            <input 
                  id=tel3 style="with: 103px" name="Ys_m_M_fax3">
            </span>传真号码</td>
            </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>手　　机:</span></div></td>
                  <td width="33%" bgcolor="#FFFFFF"> 
                    <input id=handset  name=mobile>
                    <span class="STYLE20">*</span></td>
            <td width="15%" bgcolor="#F7F7F7"><div align="right"><span 
            class=style16>邮　　编:</span></div></td>
            <td width="35%" bgcolor="#FFFFFF"><input id=post  
            name=post></td>
          </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>E-mail:</span></div></td>
                  <td width="33%" bgcolor="#FFFFFF"><span class=p141> 
                    <input  name=email>
                    </span></td>
            <td width="15%" bgcolor="#F7F7F7"><div align="right"><span 
            class=style16>网　　址:</span></div></td>
            <td width="35%" bgcolor="#FFFFFF"><span class=p141>
              <input id=url 
                   name=web value="http://">
            </span></td>
          </tr>
          <tr>
                  <td width="17%" bgcolor="#F7F7F7">
<div align="right"><span 
            class=style16>公司简介:</span></div></td>
            <td colspan="3" bgcolor="#FFFFFF"><textarea style="with: 357px; HEIGHT: 102px" rows="4" name="qyjj" cols="52"></textarea>
              <span class=red>*</span></td>
            </tr>
          <tr>
                  <td bgcolor="#F7F7F7" width="17%">
<div align="right">验 
              证 码:</div></td>
            <td colspan="3" bgcolor="#FFFFFF"><input id=txtVerify name=txtVerify  type="text" size="4">
              <script>
                    document.write("<img id=\"vcodeimage\" src=\"GetCode.asp\"  alt=\"验证码,看不清楚?请点击刷新验证码\" style=\"cursor : pointer;\" onclick=\"RefreshVerifyCode();\" >")

                    var obImg = document.getElementById("vcodeimage");
                    var timeoutid=null; 

                    function RefreshVerifyCode()
                    {
                      clearTimeout(timeoutid);
                      obImg.src = "GetCode.asp"
                      timeoutid = setTimeout("RefreshVerifyCode()", 240000);
                    }

                    timeoutid = setTimeout("RefreshVerifyCode()", 240000);
                  </script>
&nbsp;看不清验证码请刷新 </td>
            </tr>
          <tr>
                  <td bgcolor="#F7F7F7" width="17%">&nbsp;</td>
            <td colspan="3" bgcolor="#FFFFFF"><a href="User_Reg1.asp">查看服务条款</a></td>
            </tr>
          <tr>
            <td colspan="4" bgcolor="#FFFFFF"><table border="0" cellpadding="0" style="border-collapse: collapse" with="768" id="AutoNumber6" height="40" cellspacing="0" align="center">
              <tr>
                <td><p align="center">
                    <input class=button04 type="submit" value="我已经阅读并同意接受服务条款,创建我的诚信物流网会员帐号" name="submit" >
                </p></td>
              </tr>
            </table></td>
            </tr>
          
		  <input type=hidden 
value=name=which name="hidden">
          <input name=provincetwo 
type=hidden>
        </table>
        </form>
    </td>
  </tr>
</table>
</td>
  </tr>
</table>
<!--#include file="../inc/down.htm"-->
</body>
</html>
<%end sub%>