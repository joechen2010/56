<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/config.asp"-->
<%
dim RegUserName,txtVerify
RegUserName=request.form("userid")
session("RegUserName")=RegUserName

if RegUserName="" or strLength(RegUserName)>20 or strLength(RegUserName)<4 then
	founderr=true
	errmsg=errmsg & "请输入用户名(不能大于20小于5)\n"
else
  	if Instr(RegUserName,"<")>0 or Instr(RegUserName,">")>0 or Instr(RegUserName,"=")>0 or Instr(RegUserName,"%")>0 or Instr(RegUserName,chr(32))>0 or Instr(RegUserName,"?")>0 or Instr(RegUserName,"&")>0 or Instr(RegUserName,";")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,"'")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,chr(34))>0 or Instr(RegUserName,chr(9))>0 or Instr(RegUserName,"")>0 or Instr(RegUserName,"$")>0 then
		errmsg=errmsg+"用户名中含有非法字符\n"
		founderr=true
	end if
end if

if founderr=false then
	dim sqlReg,rsReg,sqladminname,rsadminname

	sqladminname="select * from Manage_user where UserName='" & RegUserName & "'"
	set rsadminname=conn.execute(sqladminname)
	
	if not(rsadminname.bof and rsadminname.eof) then
		founderr=true
		errmsg=errmsg & "朋友，想冒充管理员？换一个用户名吧老大！\n"
		session("RegUserName")=""
	else
	
    set rs=server.CreateObject ("adodb.recordset")
      sql="select * from qyml where user='"&RegUserName&"'"
      rs.open sql,conn,1,3
      if not (rs.eof and rs.bof) then
           founderr=true
		   errmsg=errmsg & "你注册的用户已经存在！请换一个用户名再试试！\n"
		   session("RegUserName")=""
      end if
	  rs.close
	  set rs=nothing
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
if (document.UserReg.question.value =="") 
{ 
alert("请输入密码提示问题！"); 
document.UserReg.question.focus(); 
return (false); 
} 
if (document.UserReg.answer.value =="") 
{ 
alert("请输入密码提示答案！"); 
document.UserReg.answer.focus(); 
return (false); 
} 

if (document.UserReg.qymc.value =="") 
{ 
alert("请输入公司名称！"); 
document.UserReg.qymc.focus(); 
return (false); 
} 
if (document.UserReg.jyfw.value =="")
{ 
alert("请输入您企业的经营和服务情况！"); 
document.UserReg.jyfw.focus(); 
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

var regexp = /^[0123456789]{1,6}$/g; 
if (regexp.test(document.UserReg.post.value)!=1)
{
alert("请输入正确的邮政编码！");
document.UserReg.post.focus(); 
document.UserReg.post.select();
return (false);
}
var filter2=/^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
if (!filter2.test(document.UserReg.email.value)) { 
alert("邮件地址不正确,请重新填写！"); 
document.UserReg.email.focus();
document.UserReg.email.select();
return false; 
} 
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
<head>
<LINK href="images/css.css" type="text/css" rel="stylesheet">
<style type="text/css">.p14 { FONT-SIZE: 14px; COLOR: #000000; LINE-HEIGHT: 130% }
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
<TABLE cellSpacing=0 cellPadding=0 width=776 align=center border=0>
  <TBODY> 
  <TR> 
    <TD vAlign=top height=323> 
      <TABLE style="BORDER-BOTTOM: #5286b5 1px solid" height=70 cellSpacing=0 
      cellPadding=0 width=776 border=0>
        <TBODY> 
        <TR> 
          <TD width=265 height=64><A href="../"><IMG height=65 
            alt=返回首页 src="images/logo.gif" width=265 align=bottom 
            border=0></A></TD>
          <TD vAlign=bottom align=middle width=511> 
            <TABLE height=50 cellSpacing=0 cellPadding=0 width=489 border=0>
              <TBODY> 
              <TR> 
                <TD width=180><IMG height=50 src="images/title_re1.gif" 
                  width=180></TD>
                <TD width=180><IMG height=50 src="images/title_re2.gif" 
                  width=180></TD>
                <TD width=129><IMG height=50 src="images/title_re3.gif" 
                  width=129></TD>
              </TR>
              </TBODY>
            </TABLE>
          </TD>
        </TR>
        </TBODY>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
        <TBODY> 
        <TR> 
          <TD height=4></TD>
        </TR>
        </TBODY>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=776 border=0>
        <TBODY> 
        <TR> 
          <TD> 
            <TABLE cellSpacing=0 cellPadding=10 width="100%" border=0>
              <TBODY> 
              <TR> 
                <TD vAlign=top align=middle width="12%" height=30><IMG 
                  height=37 src="images/e_arrow.gif" width=64></TD>
                <TD align=right width="88%"><IMG height=30 
                  src="images/htxt_01.gif" 
        width=500></TD>
              </TR>
              </TBODY>
            </TABLE>
          </TD>
        </TR>
        </TBODY>
      </TABLE>
      <table border="0" bordercolor="#111111" width="100%" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="100%"> 
            <FORM method="POST" action="User_RegOK.asp" name="UserReg" onSubmit="return FormCheck();" >
			    <input type="hidden" name="userstate" value="new">
                <input type="hidden" name="_fom_grp_" value="Ys"/>
                <input type="hidden" name="M" value="_m_M_"/>
              <table height=51 cellspacing=0 cellpadding=0 width=760 align=center 
      border=0>
                <tbody> 
                <tr> 
                  <td width=257 rowspan=2><img height=51 src="images/dot_101.jpg" 
            width=257></td>
                  <td align=right height=22><span class=S>注册会员中有任何问题请拨打:<span 
            class=style17>0574-63809664</span></span></td>
                </tr>
                <tr> 
                  <td background=images/dot_102.jpg 
      height=29>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <TABLE border=0 bordercolor="#3399FF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-collapse: collapse" cellspacing="0" cellpadding="0" align="center" width="760">
                <tr> 
                  <TD align=right width="200" bgcolor="#f7f7f7"><SPAN 
            class=style16>会员用户名:</SPAN></TD>
                  <TD width="13">&nbsp;</TD>
                  <TD width="547"><SPAN class=p14> 
                    <INPUT class=f11 maxLength=20 name="Ys_m_M_user" size="20" value="<%=session("RegUserName")%>" readonly>
                    </SPAN> <font color="#666666"><span class=red>*</span> <span class=huang12>5-20个字符（<font face="Arial">A-Z</font>， 
                    <font face="Arial"> a-z</font>， <font face="Arial"> 0-9</font>），建议您取一个便于输入和记忆的登录名</span></font></TD>
                </tr>
                <TR vAlign="middle" background="images/dot3.gif"> 
                  <TD align=right height="1" colspan="3"></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="0" bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>密 码：</SPAN></TD>
                  <TD height="13" width="13">&nbsp;</TD>
                  <TD height="13" width="547"> 
                    <INPUT class=f11 maxLength=20 name=pass type=password size="20">
                    <span class=p14><span class=p14> </span></span><span class=red>*</span> 
                    <span class=huang12>（4-20位，请使用英文(a-z)、数字(0-9)注意区分大小写）</span></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="1" colspan="3" background="images/dot3.gif"></TD>
                <TR vAlign="middle"> 
                  <TD align=right height="28" bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>确认密码：</SPAN></TD>
                  <TD height="28" width="13">&nbsp;</TD>
                  <TD height="28" width="547"> 
                    <INPUT class=f11 maxLength=20 name=confirmPassword type=password size="20">
                    <span class=p14><span class=p14> </span></span><span class=red>*</span> 
                    <span class=huang12>请重复一次您上面设置的密码</span></TD>
                <TR vAlign="middle" > 
                  <TD align=right height="1" colspan="3" background="images/dot3.gif"></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>您的姓名：</SPAN></TD>
                  <TD width="13">&nbsp;</TD>
                  <TD height="30" width="547"> 
                    <input class=f11 maxlength=100 name=name size="20">
                    <font color=#990000> </font><span class=red>*</span> <span 
            class=huang12>您的真实姓名</span></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="1" colspan="3" background="images/dot3.gif"></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>性别：</SPAN></TD>
                  <TD width="13">&nbsp;</TD>
                  <TD height="30" width="547"> 
                    <input class=f11 name=ch type=radio value="先生" checked>
                    先生&nbsp;&nbsp; 
                    <input class=f11 name=ch type=radio value="女士">
                    女士<span class=red></span></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="1" colspan="3" background="images/dot3.gif"></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>密码提示问题：</SPAN></TD>
                  <TD height="49" width="13">&nbsp;</TD>
                  <TD height="49" width="547"> 
                    <select name="question" id="pass_prob">
                      <option selected value="请选择密码提示问题">请选择密码提示问题</option>
                      <option value="我养的宠物叫什么名字">我养的宠物叫什么名字</option>
                      <option value="我喜欢的人叫什么名字">我喜欢的人叫什么名字</option>
                      <option value="我喜欢的音乐">我喜欢的音乐</option>
                      <option value="我喜欢的歌曲">我喜欢的歌曲</option>
                      <option value="我喜欢的作者">我喜欢的作者</option>
                      <option value="我喜欢的书">我喜欢的书</option>
                      <option value="我喜欢的名人">我喜欢的名人</option>
                      <option value="我喜欢的名言">我喜欢的名言</option>
                    </select>
                    <span class=red>*</span> <span 
            class=huang12>请填写自己最熟悉的问题，并牢记！用于忘记密码时，重设密码。</span></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="1" colspan="3" background="images/dot3.gif"></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="27" bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>密码提示答案：</SPAN> 
                  <TD height="27" width="13">&nbsp;</TD>
                  <TD height="27" width="547">
                    <INPUT class=f11 maxLength=20 name=answer size="20">
                    <span class=red>*</span> <span 
            class=huang12>请牢记这个答案，以便密码丢失时回答系统的提问。如果您的回答正确，系统就会自动把密码显示给您。</span></TD>
                </TR>
                <INPUT type=hidden 
value=name=which>
                <INPUT name=provincetwo 
type=hidden>
              </TABLE>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td width=257 rowspan=2><img height=53 src="images/dot_201.jpg" 
            width=322></td>
                  <td height=24><img height=1 src="images/dot.gif" width=1></td>
                </tr>
                <tr> 
                  <td background=images/dot_102.jpg 
      height=29>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <TABLE border=0 bordercolor="#3399FF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-collapse: collapse" cellspacing="0" cellpadding="0" align="center" width="760">
                <tr> 
                  <TD align=right height="18" bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>公司名称：</SPAN></TD>
                  <TD height="18" width="14">&nbsp;</TD>
                  <TD height="18"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=280><span class=p14><font face=宋体><span 
                  class=p14><font face=宋体> 
                          <input style="WIDTH: 214px" 
                  name=qymc>
                          </font></span></font></span></td>
                        <td height=30><span class=red>*</span> <span class=huang12>请填写在工商局注册的公司/商号全称</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="18" bgcolor="#f7f7f7"><SPAN 
            class=style16>产业种类：</SPAN></TD>
                  <TD height="18">&nbsp;</TD>
                  <TD height="30"> 
                    <%
        sql = "select * from Class_1"
        rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	response.write "请先添加栏目。"
	response.end
	else
%>
                    <select name="SortID">
                      <option selected value="<%=trim(rs("SortID"))%>"><%=trim(rs("sort"))%></option>
                      <%      dim selclass
         selclass=rs("SortID")
        rs.movenext
        do while not rs.eof
%>
                      <option value="<%=trim(rs("SortID"))%>"><%=trim(rs("sort"))%></option>
                      <%
        rs.movenext
        loop
	end if
        rs.close
%>
                    </select>
                    <font color=#FF0000 size=2> *</font></TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="18" bgcolor="#f7f7f7"><SPAN 
            class=style16>经营模式：</SPAN></TD>
                  <TD height="18">&nbsp;</TD>
                  <TD height="30"><font color="#000080"> 
                    <input type="radio" name="qylb" value="制造商" checked>
                    制造商 
                    <input type="radio" name="qylb" value="采购商">
                    采购商 
                    <input type="radio" name="qylb" value="经销商">
                    经销商 <span class=p14><span class=S><span 
            class=red>*</span> </span></span><span 
            class=huang12>选择主营行业</span></font></TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="18" bgcolor="#f7f7f7"><SPAN 
            class=style16>企业性质：</SPAN></TD>
                  <TD height="18">&nbsp;</TD>
                  <TD height="30"> 
                    <select size="1" name="qyxz">
                      <option value="行政机关">行政机关</option>
                      <option value="社会团体">社会团体</option>
                      <option value="事业单位">事业单位</option>
                      <option value="国有企业">国有企业</option>
                      <option value="股份集团">股份集团</option>
                      <option value="外资企业">外资企业</option>
                      <option value="三资企业">三资企业</option>
                      <option value="集体企业">集体企业</option>
                      <option value="乡镇企业">乡镇企业</option>
                      <option value="私营企业" selected>私营企业</option>
                    </select>
                    <font color="#000080">(必填)</font> </TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="4" background="images/dot3.gif">
                </tr>
                <tr> 
                  <TD align=right height="18" bgcolor="#f7f7f7"><SPAN 
            class=style16>经营范围：</SPAN> 
                  <TD> 
                  <TD height="18" colspan="2"> 
                    <textarea style="WIDTH: 357px; HEIGHT: 102px" rows="2" name="jyfw" cols="56"></textarea>
                    <font color="#FF0000"> </font><font color="#000080">(必填)</font> 
                  </TD>
                </tr>
                <INPUT type=hidden 
value=name=which>
                <INPUT name=provincetwo 
type=hidden>
              </TABLE>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td width=257 rowspan=2><img height=51 src="images/dot_301.jpg" 
            width=260></td>
                  <td height=22><img height=1 src="images/dot.gif" width=1></td>
                </tr>
                <tr> 
                  <td background=images/dot_102.jpg 
      height=29>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <TABLE border=0 bordercolor="#3399FF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-collapse: collapse" cellspacing="0" cellpadding="0" align="center" width="760">
                <tr> 
                  <TD align=right bgcolor="#f7f7f7" width="200"> 
                    <p align="right"><span 
            class=style16>公司所在地:</span> 
                  </TD>
                  <TD height="18" width="14">&nbsp;</TD>
                  <TD height="18"> 
                    <table id=Table2 cellspacing=0 cellpadding=0 width="98%" border=0>
                      <tbody> 
                      <tr> 
                        <td width="56%"> 
                          <select id="s1" name="Ys_m_M_province">
                            <option value="省份">省份</option>
                          </select>
                          <select id="s2" name="Ys_m_M_city">
                            <option>地级市</option>
                          </select>
                          <select id="s3" name="Ys_m_M_area">
                            <option>市、县级市、县</option>
                          </select>
                        </td>
                        <td class=S width="44%" height="30"><span class=red>*</span> 
                          <span 
                  class=huang12>你按菜单选择</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="27" bgcolor="#f7f7f7"><SPAN 
            class=style16>公司地址：</SPAN></TD>
                  <TD height="27">&nbsp;</TD>
                  <TD height="27"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=250> 
                          <input id=address style="WIDTH: 214px" 
                  name="Ys_m_M_address">
                          &nbsp; </td>
                        <td width=290><span class=red>*</span> <span 
                  class=huang12>只需填写街道、门牌号</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="30" bgcolor="#f7f7f7"><SPAN 
            class=style16>固定电话：</SPAN></TD>
                  <TD height="30">&nbsp;</TD>
                  <TD height="30"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=66 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel1 style="WIDTH: 37px" value=+86 name="Ys_m_M_phone1">
                          </span></span>- </td>
                        <td width=99 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel2 style="WIDTH: 53px" name="Ys_m_M_phone2">
                          </span></span>-</td>
                        <td width=82 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel3 style="WIDTH: 103px" name="Ys_m_M_phone3">
                          </span></span></td>
                        <td width=293 height=30><span class=red>*</span> <span 
                  class=huang12>国家代码-区号-电话号码</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="30" bgcolor="#f7f7f7"><SPAN 
            class=style16>传真：</SPAN></TD>
                  <TD height="30">&nbsp;</TD>
                  <TD height="30"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=66 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel1 style="WIDTH: 37px" value=+86 name="Ys_m_M_fax1">
                          </span></span>- </td>
                        <td width=99 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel2 style="WIDTH: 53px" name="Ys_m_M_fax2">
                          </span></span>-</td>
                        <td width=82 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel3 style="WIDTH: 103px" name="Ys_m_M_fax3">
                          </span></span></td>
                        <td width=293 height=30><span 
                class=huang12>传真号码</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" background="images/dot3.gif" colspan="3"></TD>
                </tr>
                <tr> 
                  <TD align=right height="27" bgcolor="#f7f7f7"><SPAN 
            class=style16>手机：</SPAN></TD>
                  <TD height="27">&nbsp;</TD>
                  <TD height="30"> 
                    <input id=handset style="WIDTH: 214px" 
            name=mobile>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" background="images/dot3.gif" colspan="3"></TD>
                </tr>
                <tr> 
                  <TD align=right height="29" bgcolor="#f7f7f7"><SPAN 
            class=style16>邮编：</SPAN></TD>
                  <TD height="29">&nbsp;</TD>
                  <TD height="30"> 
                    <input id=post style="WIDTH: 214px" 
            name=post>
                    &nbsp;<span class=huang12>邮编号码</span> </TD>
                </tr>
                <tr> 
                  <TD height="1" background="images/dot3.gif" colspan="3"></TD>
                </tr>
                <tr> 
                  <TD align=right height="29" bgcolor="#f7f7f7"><SPAN 
            class=style16>E-mail：</SPAN></TD>
                  <TD height="29">&nbsp;</TD>
                  <TD height="29"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=280><span class=p14> 
                          <input style="WIDTH: 214px" name=email>
                          </span></td>
                        <td height=30><span class=red>*</span> <span 
                  class=huang12>您的电子邮箱地址,例：&nbsp; info@cx56w.com</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" background="images/dot3.gif" colspan="3"> 
                </tr>
                <tr> 
                  <TD align=right height="29" bgcolor="#f7f7f7"><SPAN 
            class=style16>网址：</SPAN></TD>
                  <TD height="29">&nbsp;</TD>
                  <TD height="29"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=247><span class=p14> 
                          <input id=url 
                  style="WIDTH: 214px" name=web value="http://">
                          </span></td>
                        <td class=huang12 width=293 
              height=30>输入您原有网站网址，在诚信物流网得到一并推广</td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" background="images/dot3.gif" colspan="3"></TD>
                </tr>
                <tr> 
                  <TD align=right height="29" bgcolor="#f7f7f7"><SPAN 
            class=style16>公司简介:</SPAN></TD>
                  <TD height="29">&nbsp;</TD>
                  <TD height="29"> 
                    <textarea style="WIDTH: 357px; HEIGHT: 102px" rows="5" name="qyjj" cols="56"></textarea>
                    <span 
            class=red>*</span> <span class=huang12>公司简介字数不得少于25个</span></TD>
                </tr>
                <tr>
                  <TD align=right height="29" bgcolor="#f7f7f7">验证码：</TD>
                  <TD height="29">&nbsp;</TD>
                  <TD height="29">
<input id=txtVerify name=txtVerify  type="text" size="4">
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
&nbsp;看不清验证码请刷新 </TD>
                </tr>
              </TABLE>
              <table border="0" cellpadding="0" style="border-collapse: collapse" width="768" id="AutoNumber6" height="88" cellspacing="0">
                <tr> 
                  <td width="100%" height="40" align="center">【 <a class=red 
            href="Item.asp" 
            target=_blank>点此阅读诚信物流网服务条款</a> 】</td>
                </tr>
                <tr> 
                  <td width="100%"> 
                    <p align="center"> 
                      <input class=button04 type="submit" value="我已经阅读并同意接受服务条款,创建我的诚信物流网会员帐号" >
                    </p>
                  </td>
                </tr>
              </table>
            </FORM>
          </td>
        </tr>
      </table>
    </TD>
  </TR>
  </TBODY>
</TABLE>
<!--#include file="../inc/down.htm"-->
</body>
</html>
<%end sub%>