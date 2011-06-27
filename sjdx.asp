<!--#include file="Inc/Conn2.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
const MaxNumber=50
const MaxNum=8
const MaxLen=30
dim rs_sel,sel1_total,rs_buy,buy_total,rs_p,p_total
set rs_sel=conn.execute("select Count(*) from Info where InfoType='供应'")
sell_total=rs_sel(0)
rs_sel.close
set rs_sel=nothing

set rs_buy=conn.execute("select Count(*) from Info where InfoType='采购'")
buy_total=rs_buy(0)
rs_buy.close
set rs_buy=nothing

set rs_s_o=conn.execute("select Count(*) from offer where OfferType='供应'")
sell_o_total=rs_s_o(0)
rs_s_o.close
set rs_s_o=nothing

set rs_b_o=conn.execute("select Count(*) from offer where OfferType='采购'")
buy_o_total=rs_b_o(0)
rs_b_o.close
set rs_b_o=nothing

set rs_p=conn.execute("select Count(*) from qyml")
p_total=rs_p(0)
rs_p.close
set rs_p=nothing
%>
<%typeselect=request("type")
%>
<%
function title_left(str,num)
if len(str)>num then
response.write left(str,num)&""
else
response.write str
end if
end function
%>
<HTML>
<HEAD>
<TITLE>168货运信息网--全国最专业最诚信的物流信息交易平台|物流,物流网,中国物流网,物流信息网</TITLE>
<meta name=keywords content="物流,物流网,中国物流网,168货运信息网">
<meta name=description content="物流-168货运信息网是大型关于物流资讯和物流产品应用,产品厂家介绍的网络媒体.欢迎成为168货运信息网注册会员提供更多服务!Find the logistics products and services in this website!">
<LINK href="images/css.css" type=text/css rel=stylesheet>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content=. name=description>
<META  content=，>
<SCRIPT language=JavaScript type=text/JavaScript>
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
function openwindow(url,width,height) { 	left1 =0; 	top1 = 0; 	window.open(url,"","width=" + width + ",height=" + height + ",toolbar=no,menubar=no,scrollbars=yes,location=no,status=yes,left=" + left1.toString() + ",top=" + top1.toString()); }
//resizable=yes, 最大化
//-->
</SCRIPT>
<script language=JavaScript src="inc/p_c_a.js"></script>
<META content="MSHTML 6.00.2462.0" name=GENERATOR>
<style type="text/css">
<!--
.STYLE1 {color: #FFFFFF}
-->
</style>
</HEAD>
<BODY topMargin=0>
<!--#include virtual="sjtop.asp" -->
<table cellspacing=0 cellpadding=0 width=768 align=center border=0>
  <tbody> 
  <tr> 
    <td background=images/backdropi.gif bgcolor=#ffffff> 
      <table cellspacing=0 cellpadding=0 width="96%" align=center border=0>
        <tbody> 
        <tr> 
          <td><img height=100 src="images/sjbt.gif" 
        width=760></td>
        </tr>
        <tr> 
          <td background=images/fz02.gif> 
            <div align=center> 
              <table style="TEXT-INDENT: 25px; LINE-HEIGHT: 150%" cellspacing=1 
            cellpadding=7 width="96%" bgcolor=#d9d9d9 border=0>
                <tbody> 
                <tr bgcolor=#eeffdb> 
                  <td valign=top width="30%" 
                  height=88>168货运信息网开通了支持手机,小灵通短信<br>
                    双向互动,跨地区全国支持的短信特服号． 
                    <p>如果您是168货运信息网的会员，并且在注册时填写了正确的手机号或小灵通号，并且在您的会员帐户里预存了短信资费，就可以通过短信查询配货信息．<br>
                      <img 
                  height=77 src="images/fsgz.gif" width=214></p>
                    <p>当您在168货运信息网总站或分站或中华物流通中发布了货源或空车信息，在有效期内如果其他会员发布了一条和您的信息相匹配的信息，则你的手机会收到这条信息． 
                    <p>例如，您今天在网站上发布了一条从长春到北京的货源信息，有效期为3天，那么在今后的3天内，如果有其他会员发布了从长春到北京的车源信息，则168货运信息网会自动将这条信息发到你的手机上，供您参考． 
                      <br>
                      <font color="#000080">天津到重庆市区有车25吨宏达配货站 张宏伟 022-4724556 
                      13545238688 [短信余额288条]：168货运信息网</font> </p>
                  </td>
                  <td valign=top> 
                    <p><strong>查询车源编辑短信</strong><br>
                      出发城市区号（如010）星号（＊）到达城市区号（如021）C <br>
                      例如，您有货要从北京运往上海，想查询北京到上海的车源，则发送如下短信：<font color=#ff0000><b> 
                      <font size=2>010*021C</font></b></font> 
                    <p><strong>查询货源编辑短信</strong><br>
                      出发城市区号（如0379）星号（＊）到达城市区号（如0431）H&nbsp; <br>
                      例如，您有车，想查询从郑州运往长春的货源，则发送如下短信：<font color=#ff0000><b><span 
                  class=style10><font 
                  size=2>0371*0431H</font></span></b></font><font size=2> </font><br>
                      移动用户发送到 <b><font size="2" color="#800080">620127870290</font></b><br>
                      联通用户发送到 <b><font size="2" color="#800080">620127870290</font></b><br>
                    </p>
                    <p>如果有满足查询条件,并且在有效期内的信息，则168货运信息网将发送类似与如下短信到你的手机/小灵通上： <br>
                      <font 
                  color=#000080>北京市区到上海市区有车30吨宏达物流 张先生01062500559 13114861250[还有2条信息] 
                      [短信余额288条]：168货运信息网<br>
                      杭州到武汉重货25吨九龙物流 郑明0571-85953056 13642182516急发[还有5条信息][短信余额267条]：168货运信息网</font><br>
                      <font 
                  color=#000080>　　[还有5条信息]</font> 表明还有5条满足查询条件的信息，您可以再次发送同样的查询短信，168货运信息网将发送下一条信息给你． 
                      <br>
                      　　如果从您手机发出的查询短信格式不对，或没有满足查询条件的信息，则168货运信息网不发送，以节省您在168货运信息网预存的短信资费．</p>
                    </td>
                </tr>
                <tr bgcolor=#eeffdb> 
                  <td width="30%" height=88><img height=77 
                  src="images/zfbz.gif" width=214></td>
                  <td valign=top>从您的手机/小灵通发送查询短信时，像你发送其它短信一样，电信部门将扣除0.1元． <br>
                    　　从168货运信息网向您的手机/小灵通发送配货信息时，168货运信息网从您的会员帐户里预存的短信资费中扣除0.1元． 
                    <br>
                    &nbsp;&nbsp;&nbsp; 您可以随时在会员中心或中华物流通里查看168货运信息网给你发送配货短信的记录．</td>
                </tr>
                </tbody> 
              </table>
            </div>
          </td>
        </tr>
        <tr> 
          <td><img height=22 src="images/fz03.gif" 
        width=760></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<tr> 
  <td align="center"> 
    <!--#include file="sjbottom.html" -->
    <!-- Live800在线客服图标:cx56w1[浮动型] 开始-->
    <script language="javascript" src="http://chat.live800.com:8080/live800/chatClient/floatButton.js?jid=9756223015&companyID=32913&configID=24234&codeType=custom"></script>
    <!-- Live800在线客服图标:cx56w1[浮动型] 结束-->
    <!-- Live800在线客服功能代码 开始-->
    <script language="javascript" src="http://chat.live800.com:8080/live800/chatClient/monitor.js?jid=9756223015&companyID=32913&configID=23583&codeType=custom"></script>
    <!-- Live800在线客服功能代码 结束-->
</BODY>
</HTML>
