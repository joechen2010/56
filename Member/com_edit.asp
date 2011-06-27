<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<%
set rs=conn.execute("select * from qyml where user='"&session("user")&"'")
if rs.eof and rs.bof then
    response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	response.end
else
qyjj=replace(rs("qyjj"),"&nbsp;",chr(32))
qyjj=replace(qyjj,"<br>",chr(13))
qyjj=replace(qyjj,"<BR>",chr(13))

%>
<HTML>
<HEAD>
<TITLE>诚信物流网 会员控制中心 - 物流商人的网上家园</TITLE>
<meta http-equiv="Content-Language" content="zh-cn"> 
<META qyjj=zh-cn http-equiv=qyjj-Language> 
<META qyjj="text/html; charset=gb2312" http-equiv=qyjj-Type> 
<META qyjj="Microsoft FrontPage 5.0" name=GENERATOR content="Microsoft FrontPage 5.0">
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
.STYLE3 {color: #FF0000}
-->
</style>
<script language=JavaScript src="../inc/p_c_a.js"></script>
<SCRIPT language=javascript src="../../login/images/location.js"></SCRIPT> 
<SCRIPT language=JavaScript type=text/javascript>
//----添加到选择框
function addit(from, to) {
var src = document.getElementById(from).options;
var dst = document.getElementById(to).options;

var selected_value = [];
var selected_text = [];

//获取已选择的类别
for(var i = 0; i < dst.length; i++) {
  selected_value[i] = dst[i].value;
  selected_text[i] = dst[i].text;
}

// 获取源类别
for(var i = 1; i < src.length; i++) {
  if(src[i].selected) {
		var exists = false;
	
		for(var j = 1; j < dst.length; j++) {
		  if(dst[j].value == src[i].value) {
			exists = true;
			break;
		  }
		}
		
		if(exists==false) {
		  dst.length++;
		  selected_value[dst.length-1] = src[i].value;
		  selected_text[dst.length-1] = src[i].text;
		}
  }
}

	//填充已选择框
	for(var j = 0; j < selected_value.length; j++) {
	  if( selected_value[j] == "" ) continue;
	  dst[j].value = selected_value[j];
	  dst[j].text = selected_text[j]
	}
}

//删除已选择框内的类别
function delit(dst) {
	var menu=document.getElementById(dst).options;
	for(var i = 1; i < menu.length; i++) {
	  if(menu[i].selected) menu[i--] = null;
	}
}

//----设置隐藏表单域
function tofull(){
	
	
<%
set rs_s2=conn.execute("select * from Class_1 Order By SortID Asc")
if rs_s2.eof and rs_s2.bof then
'
else
do while not rs_s2.eof
%>  
    var tempstr<%=rs_s2("SortID")%> = "";
	var bearings<%=rs_s2("SortID")%> = document.getElementById("bearings<%=rs_s2("SortID")%>").options;
	var fullbearings<%=rs_s2("SortID")%> = document.getElementById("fullbearings<%=rs_s2("SortID")%>");
	for(var i = 1;i < bearings<%=rs_s2("SortID")%>.length;i++){
		tempstr<%=rs_s2("SortID")%> = tempstr<%=rs_s2("SortID")%>+","+bearings<%=rs_s2("SortID")%>[i].value;
	}
	fullbearings<%=rs_s2("SortID")%>.value = tempstr<%=rs_s2("SortID")%>.substr(1,tempstr<%=rs_s2("SortID")%>.length);
<%
rs_s2.movenext
loop
end if
rs_s2.close
set rs_s2=nothing
%>
}
//-----计算文本长度（区分英文和汉字）
function countlength(str) 
{	
	var strlength = 0;
    var pattern = /[\u4E00-\u9FA5]|[\uFE30-\uFFA0]/; 
	for (var i = 0;i < str.length;i++){
		if (pattern.test(str.charAt(i))) {
			strlength = strlength + 1;
		}
		else {
			strlength = strlength + 0.5;
		}
	}
	return strlength;
}
//去两端空格
function trim(str) {
	regExp1 = /^ */;
	regExp2 = / *$/;
	return str.replace(regExp1,'').replace(regExp2,'');
} 

//----提交验证函数
function checksubmit() {
//验证公司名
	var ename = trim(document.getElementById("qymc").value);	
	if (ename==""){
		alert("公司名称必填，请填写贵公司在工商部门登记的全称!");
		document.getElementById("qymc").focus();
		return false;
	}
	else if(countlength(ename)>80){
		alert("公司名称不能超过80个字符,如果超过,请尽量缩减!");
		document.getElementById("qymc").focus();
		return false;
	}

//验证企业性质
	var qyxz = document.getElementById("qyxz").value;
	if (qyxz==""){
		alert("请选择企业性质!");
		document.getElementById("contact").focus();
		return false;
	}
	
//验证法人代表
	var legalperson = document.getElementById("frdb").value;
	if(countlength(legalperson)>0){
		if ((countlength(legalperson)>4)||(countlength(legalperson)<2)){
			alert("法人代表长度在2-4个汉字之间,或者请留空!");
			document.getElementById("frdb").focus();
			return false;
		}
	}
//验证注册资金

//验证产业类型
	var industry1 = document.getElementById("qylb1").checked;
	var industry2 = document.getElementById("qylb2").checked;
	var industry3 = document.getElementById("qylb3").checked;
	if ((industry1||industry2||industry3)==false){
		alert("请至少选择一项产业!");
		document.getElementById("qylb1").focus();
		return false;
	}
	
//验证荣誉
	var jygs = document.getElementById("jygs").value;
	if(countlength(jygs)>200){
		alert("公司荣誉请控制在200个汉字内，当前为 "+countlength(jygs)+" 个");
		document.getElementById("jygs").focus();
		return false;
	}	
//验证通过，提交表单
	document.getElementById("form1").submit();
}
</script>
</head>
<body> 
<!--#include file="head.htm"-->
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
      <table width="534"  border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 修改公司资料</td>
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
                <td><table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="images/title_2.jpg" bgcolor="#FFFFFF" >
                        <div align="center"><font color="#FFFFFF">修改公司资料</font></div>
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
                              <td align="left"><div id="TipAll" >
                               请输入所有标有*的项目。
                              </div></td>
                            </tr>
                           </table>
<table width="100%"   border="0" cellpadding="0" cellspacing="0">
                    
					<tr>
                      <td valign="top">
                        <table  border="0" cellpadding="5" cellspacing="1" bgcolor="#B1D4F2" >
                          <FORM method="POST" action="com_edit_save.asp" name="Form1">
                            <TBODY> 
                            <tr> 
                              <TD align=left height="25" colspan="2" class="white" background="images/title_bg.gif">■ 
                                修改公司资料</TD>
                            </TR>
                            <tr bgcolor="#EDF6FF"> 
                              <TD align=right height="24" width="150"> <font color="#FF0000">*</font> 
                                公司名称：</TD>
                              <TD height="24" bgcolor="#EDF6FF"> 
                                <INPUT class="box" maxLength=80 name=qymc size=36 value=<%=rs("qymc")%>>                              </TD>
                            </tr>
                            
                            <tr bgcolor="#EDF6FF"> 
                              <TD width="150" height="24" align=right bgcolor="#EDF6FF">产业种类：</TD>
                              <TD height="24" bgcolor="#EDF6FF">
<%set rs7=conn.execute("select * from Class_1 order by Sortid asc")

	if rs7.eof and rs7.bof then
	response.write "请先添加栏目。"
	response.end
	else
%>
              <select name="SortID">
                <option selected value="<%=trim(rs("qylb"))%>"><%=trim(rs("qylb"))%></option>
                <%do while not rs7.eof%>
                <option value="<%=trim(rs7("sort"))%>"><%=trim(rs7("sort"))%></option>
                <%
        rs7.movenext
        loop
	end if
        rs7.close
		set rs7=nothing
%>
              </select>							  

							 </TD>
                            </tr>
                            <tr bgcolor="#EDF6FF">
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span class="style16">主要司机:</span></div></td>
            <td width="35%" bgcolor="#EDF6FF"><input type="text" name="siji" value="<%=rs("siji")%>"></td>
                            </tr>
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>公司所在地:</span></div></td>
            <td  bgcolor="#EDF6FF"><select id="s1" name="province">
              <option value="省份">省份</option>
            </select>
              <select id="s2" name="city">
                <option>地级市</option>
              </select>
              <select id="s3" name="area">
                <option>市、县级市、县</option>
              </select>
              <span class=red>*</span></td>
            </tr>
						<SCRIPT language=javascript>
            setup();
          </SCRIPT>      
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>公司地址:</span></div></td>
            <td bgcolor="#EDF6FF"><input id=address  name="address" value="<%=rs("address")%>">
&nbsp; <span class=red>*</span> 只需填写街道、门牌号</td>
            </tr>	
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>固定电话:</span></div></td>
            <td bgcolor="#EDF6FF"><span class=p141>
              <input id=Ys_m_M_phone1 style="with: 37px" name="phone1" type="hidden" value="<%=rs("phone1")%>">
            </span><span class=p141>
            <input 
                  id=Ys_m_M_phone2 style="with: 53px" name="phone2" size="4" value="<%=rs("phone2")%>">
            </span>-<span class=p141>
            <input 
                  id=Ys_m_M_phone3 style="with: 103px" name="phone3" value="<%=rs("phone3")%>">
            </span><span class=red>* </span>区号-电话号码</td>
            </tr>
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>传　　真:</span></div></td>
            <td bgcolor="#EDF6FF"><span class=p141>
              <input id=tel1 style="with: 37px" name="fax1" type="hidden" value="<%=rs("fax1")%>">
            </span><span class=p141>
            <input id=tel2 style="with: 53px" name="fax2" size="4" value="<%=rs("fax2")%>">
            </span>-<span class=p141>
            <input id=tel3 style="with: 103px" name="fax3" value="<%=rs("fax3")%>">
            </span>传真号码</td>
            </tr>			
          <tr>
            <td width="150" bgcolor="#EDF6FF" align="right">手　机:</td>
            <td bgcolor="#EDF6FF"><input id=handset  name=mobile value="<%=rs("mobile")%>"></td>
            </tr>
          <tr>
            <td width="150" bgcolor="#EDF6FF" align="right">邮　　编:</td>
            <td bgcolor="#EDF6FF"><input id=post name=post value="<%=rs("post")%>"></td>
            </tr>
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>E-mail:</span></div></td>
            <td width="35%" bgcolor="#EDF6FF"><span class=p141>
              <input  name=email value="<%=rs("email")%>">
            </span></td>
            </tr>
          <tr>
            <td width="150" bgcolor="#EDF6FF"><div align="right"><span 
            class=style16>网　　址:</span></div></td>
            <td width="35%" bgcolor="#EDF6FF"><span class=p141>
              <input id=url name=web value="<%=rs("web")%>">
            </span></td>
            </tr>									

                            <tr bgcolor="#EDF6FF"> 
                              <TD align=right width="150"><font color="#FF0000">*</font> 
                                公司简介：</TD>
                              <TD bgcolor="#EDF6FF">
                                <textarea name="content" cols="60" rows="10" id="content"><%=dvHTMLCode(rs("qyjj"))%></textarea>                              </TD>
                            </tr>
                            <tr bgcolor="#EDF6FF"> 
                              <TD align=right height="36" colspan="2"> 
                                <p align="center"> 
                                  <input type="submit" value=" 确定修改 " class="btsub">
                                  &nbsp;&nbsp;&nbsp; 
                                  <input type="button" value=" 取消返回 " onclick=javascript:history.go(-1)  class="btsub">
                              </TD>
                            </tr>
                          </FORM>
                        </table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="10"></td>
                            </tr>
                        </table>
                          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              
                            <td width="7" height="6"><img src="images/con_1.gif" width="7" height="6"></td>
                              
                            <td background="images/con_bg1.gif"></td>
                              
                            <td width="6" height="6"><img src="images/con_2.gif" width="6" height="6"></td>
                            </tr>
                            <tr> 
                              
                            <td background="images/con_bg2.gif"></td>
                              <td align="left" valign="top" bgcolor="#FEFFD2"> 
                                <table width="100%"  border="0" cellpadding="5" cellspacing="0">
                                  <tr> 
                                    <td align="left">请仔细核对您提交的信息，如果信息有误，会影响您的交易。</td>
                                  </tr>
                                  <tr>
                                    <td height="1" align="left" bgcolor="#FFCC00"></td>
                                  </tr>
                                  <tr> 
                                    <td align="left">公司联系方式与您的会员个人资料中的注册信息不冲突，最好保持一致。 
                                    </td>
                                  </tr>
                                </table> </td>
                              
                            <td background="images/con_bg4.gif"></td>
                            </tr>
                            <tr> 
                              
                            <td width="7" height="6"><img src="images/con_3.gif" width="7" height="6"></td>
                              
                            <td background="images/con_bg3.gif"></td>
                              
                            <td width="6" height="6"><img src="images/con_4.gif" width="6" height="6"></td>
                            </tr>
                          </table>
	                  </td>
					</tr>
                  </table></td>
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
</HTML>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing

sub showtypeoption(id)
dim rs_tt
ssql="select * from Class_2 Where TypeID="&cstr(id)
	set rs_tt=conn.execute(ssql)
	if rs_tt.eof and rs_tt.bof then
	    response.write "<option value=""11"">11</option>"
	else
		response.write "<option value="""&rs_tt("typeid")&""">"&rs_tt("typename")&"</option>"
	end if
	rs_tt.close
	set rs_tt=nothing
end sub
%>