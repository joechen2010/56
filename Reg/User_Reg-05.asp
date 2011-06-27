<%@language=vbscript codepage=936 %>
<%
option explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<HTML>
<HEAD>
<TITLE>会员注册确认页面-诚信物流网</TITLE>
<META content=False name=vs_snapToGrid>
<META content=True name=vs_showGrid>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META 
content=诚信物流网:全球最大的物流电子商务网站，提供物流品供求、物流品企业、物流产品、物流品项目、物流品新闻、物流品行情、物流品会展等信息的专业物流品网站。 
name=description>
<META 
content=全球最大的B2B物流平台、传承七千年物流文明精华，催生21世纪物流先锋、物流品供求、物流品、陶瓷物流、金属物流、泰蓝制品、艺术根雕、漆器雕漆、物流木雕、石雕物流、金箔制品、玉器玛瑙、红木摆件、奇石物流、象牙雕刻. 
name=keywords>
<META content=TRUE name=MSSmartTagsPreventParsing>
<META http-equiv=MSThemeCompatible content=Yes><LINK href="images/css.css" 
type=text/css rel=stylesheet>
<script language="JavaScript"> 
var isOK;
function fCheck(){
	/*something wrong*/
	/*form.UserName.trim();*/
	if( document.form.userid.value =="") {
		alert("\请输入您的用户名 !");
		document.form.userid.focus();
		return false;
	}
	if( document.form.userid.value.length<5||document.form.userid.value.length>20) {
		alert("\您的用户名长度应该在5－20个字符之间!");
		document.form.userid.focus();
		return false;
	}
	if ( fIsNumber(document.form.userid.value.charAt(0),"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")!=1 ){
		alert("\您的用户名只能以字母开头!");
		document.form.userid.focus();
		return false;
	}
	if ( fIsNumber(document.form.userid.value,"1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._-")!=1 ){
		alert("\您的用户名应该是数字、字母,不允许出现汉字等其他字符!");
		document.form.userid.focus();
		return false;		
	}
	 var mystring = "a" + document.form.userid.value;
       if (mystring.indexOf("falun")>0 ){
               alert("您使用了非法的用户名,请重新选择一个");
               return false;
       }
	return true;
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
<STYLE type=text/css>.p14 {
	FONT-SIZE: 14px; COLOR: #000000; LINE-HEIGHT: 130%
}
.S {
	FONT-SIZE: 12px
}
.style16 {
	FONT-WEIGHT: bold
}
.button {
	BORDER-RIGHT: #ffffff thin; BACKGROUND-POSITION: center center; BORDER-TOP: #ffffff thin; FONT-SIZE: 12px; BORDER-LEFT: #ffffff thin; WIDTH: 50px; COLOR: #ffffff; BORDER-BOTTOM: #ffffff thin; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 20px; BACKGROUND-COLOR: #999999
}
.button01 {
	BORDER-RIGHT: #9c9a9c 1px solid; BACKGROUND-POSITION: center center; BORDER-TOP: #9c9a9c 1px solid; FONT-SIZE: 12px; MARGIN: 1px; BORDER-LEFT: #9c9a9c 1px solid; WIDTH: 100px; COLOR: #666666; BORDER-BOTTOM: #9c9a9c 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: auto; BACKGROUND-COLOR: #cccccc
}
.button02 {
	BORDER-RIGHT: #9c9a9c 1px solid; BACKGROUND-POSITION: center center; BORDER-TOP: #9c9a9c 1px solid; FONT-SIZE: 12px; MARGIN: 1px; BORDER-LEFT: #9c9a9c 1px solid; WIDTH: 50px; COLOR: #666666; BORDER-BOTTOM: #9c9a9c 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 20px; BACKGROUND-COLOR: #cccccc
}
.button04 {
	BORDER-RIGHT: #aaaaaa 1px solid; BORDER-TOP: #fefefe 1px solid; FONT-SIZE: 12px; BACKGROUND: #ececec; BORDER-LEFT: #fefefe 1px solid; WIDTH: 380px; COLOR: #000000; PADDING-TOP: 1px; BORDER-BOTTOM: #aaaaaa 1px solid; WHITE-SPACE: normal; HEIGHT: 23px
}
.button03 {
	BORDER-RIGHT: #aaaaaa 1px solid; BORDER-TOP: #fefefe 1px solid; FONT-SIZE: 12px; BACKGROUND: #ececec; BORDER-LEFT: #fefefe 1px solid; COLOR: #000000; PADDING-TOP: 1px; BORDER-BOTTOM: #aaaaaa 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 23px
}
.style17 {
	FONT-WEIGHT: bold; COLOR: #cc3333
}
</STYLE>
<META content="MSHTML 6.00.2462.0" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginwidth="0" marginheight="0">
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
<table id=Table2 cellspacing=0 cellpadding=0 width=759 border=0 align="center">
  <tbody> 
  <tr> 
    <td valign=top height=7><img height=7 src="images/m_15.gif" 
  width=778></td>
  </tr>
  <tr> 
    <td align=middle background=images/m_13.gif height=60><img height=41 
      alt=会员注册第一步：会员服务条款 src="images/m_03.gif" width=639></td>
  </tr>
  <tr> 
    <td align=middle background=images/m_13.gif> 
      <table width="96%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td> 
            <p><br>
              <span class="p14hon">使用168货运信息网及全国各地分网所提供的服务，则要遵守如下服务条款</span><br>
              <br>
              1： 168货运信息网平台产品、品牌、VI、商标等均为诚信物流网所有，未经许可，任何单位或个人不 得滥用、非法盗用、以及任何形式的修改上述知识产权。 
              <br>
              <br>
              2： 诚信物流网站的用户应遵守中华人民共和国相关的法律和法规，遵守诚信物流网服务程序，不发表任何反动言论以及进行黄色淫秽内容的传播， 
              否则一切后果由发布人自负，168货运信息网不承担任何法律责任，诚信物流网会将发布人的一切相关信息提供给公安机关。 <br>
              <br>
              3：诚信物流网用户应提供准确的注册资料。对于用户资料不全或不准确所造成的服务中断或缺陷或引起任何损失和后果，诚信物流网不负任何责任。 
              <br>
              <br>
              4： 诚信物流网不会擅自公开用户隐私资料，但共享的注册信息，例如企业信息、个人名称、联系电话、通讯方式等可无须用户许可，可向其他正式注册用户公开共享。诚信物流网有权以电子邮件或常规邮件形式为其用户提供其他附加信息服务，且无须用户许可。 
              <br>
              <br>
              5： 诚信物流网会提供与其他国际互联网网站或资源进行链接。对于链接的其它网站或资源是否可以利用，诚信物流网不予担保，因使用或依赖上述网站或资源所生的损失或损害，网站也不负担任何责任。<br>
              <br>
              6： 诚信物流网所有内容的版权归原文作者和联盟网共同所有，任何复制、转载和再造行为均应征得原文作者和联盟网书面授权。 <br>
              <br>
              7： 诚信物流网用户发布的信息归诚信物流网所有，任何复制、转载行为都必须得到诚信物流网的书面授权，未经面授权，任何用户不得用任何形式、任何手段将任何信息转发至其他网络，否则本网有权随时终止用户的服务协议，所有已收费用概不返还， 
              并保留法律处理的权利。 <br>
              <br>
              8： 未经本网许可，任何用户不得利用本网发布除物流信息外的任何广告信息，否则有权随时暂停用户服务，严重者联盟网有权终止用户服务，所有已收费用概不返还。 
              <br>
              <br>
              9： 诚信物流网是基于互联网的公共信息交易平台，任何人不得乱发与物流行业务无关的信息，不得发布污辱信息，黄色信息以及各种法律法规不允许的，否则由此带来的后果，由发布者自负，诚信物流网不承担任何法律责任。 
              <br>
              <br>
              10： 诚信物流网只提供信息共享平台，并不参与交易过程，用户需认清交易方的真实身份，以及认真处理交易过程中的每一个环节，对于交易过程引发的一切纠纷，均由交易双方负责，168货运信息网不承担任何责任。<br>
              <br>
              11： 诚信物流网采用会员制的管理模式，用户在享受本网提供的服务前必须先付一定费用，资费标准由联盟网决定，并可根据市场情况随时调整。 
              <br>
              <br>
              12： 诚信物流网的信息采取开放的形式，游客或没有登陆的用户均可以看到信息，但信息的联系方式将被隐藏，只有正式交费用户才能获取完整的服务。 
              <br>
              <br>
              选择接受表示同意以上会员服务条款。以上协议解释权归168货运信息网。 </p>
            <p></p>
            <p></p>
          </td>
        </tr>
      </table>
      <div align=center> 
        <table cellspacing=0 width=360 border=0>
          <tbody> 
          <tr> 
            <td valign=center align=middle width=198 height=20><a 
            href="User_Reg1.asp" target=_top><img height=21 
            alt=点击进行下一步 src="images/reg_05.gif" width=56 border=0></a></td>
            <td valign=center align=middle width=197><a 
            href="User_Reg.asp" target=_top><img height=21 
            alt=返回上一页 src="images/reg_06.gif" width=56 
        border=0></a></td>
          </tr>
          </tbody> 
        </table>
      </div>
    </td>
  </tr>
  <tr> 
    <td valign=bottom><img height=7 src="images/m_16.gif" 
  width=778></td>
  </tr>
  </tbody> 
</table>
<!--#include file="../inc/down.htm"-->
</BODY></HTML>
