<!--#include file="../Inc/Conn.asp"-->
<!--#include file="../Inc/Function.asp"-->
<%
const MaxNumber=50
const MaxNum=8
const MaxLen=28
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
<HTML>
<HEAD>
<TITLE>168货运信息网--全国最专业的物流信息交易平台</TITLE>
<LINK href="../images/css.css" type=text/css rel=stylesheet>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content=诚信、诚信诚信诚信网上物流/网上物流信息/物流/车源信息/货源信息，配货/配货网,货运/货运网/货运信息网,诚信/诚信物流，诚信货运信息/诚信物流信息,诚信物流信息网/物流信息网,物流信息交易平台/诚信物流信息,物流/配货. 
name=description>
<META  content=诚信，诚信物流，诚信物流信息，诚信物流信息网，诚信，cx56w,诚信信息网，诚信配货网，诚信配货信息，诚信配货信息，网上物流/网上物流信息/物流/车源信息/货源信息，配货/配货网,货运/货运网/货运信息网,诚信/诚信物流，诚信货运信息/诚信物流信息,诚信物流信息网/物流信息网,物流信息交易平台/诚信物流信息,物流/配货,物流，物流信息，物流信息交易，物流信息交易平台，诚信物流信息，诚信物流，诚信信息交易平台，全国物流，全国物流信息，中国物流，中国物流信息，物流网，物流信息网，全国物流网，全国物流信息网，中国物流网，中国物流信息网，配货，配货信息，全国配货，全国配货信息，中国配货， 
name=keywords 中国物流，物流信息，物流软件，物流信息软件，配货站助手，车源信息，货源信息，中国物流联盟网 
高新技术，因特网服务,在线服务，上海物流,企业物流,东莞物流,厦门物流,成都物流,广东物流,浙江物流,北京物流,江苏物流,广州物流,中华大物流,中国物流在线,中国企业物流,电话号簿,电信物流,工商物流,号簿,汕头物流,短信物流,114,在线物流,物流在线,网上物流,电子物流,光盘物流,移动物流,物流查询,物流搜索,网上114,企业物流,网上推广,网络营销,中小企业,行业目录,物流数据,全球物流,物流服务,物流指南,科学,教育,文化,医药,卫生,保健,公用事业,生活服务,餐饮,娱乐,购物,旅游,居家生活用品,服装,皮革,纺织,文体休闲用品,礼品,工艺品,商务服务,广告,物流,印刷,包装,造纸,工程,建筑,房地产,装潢,计算机,互联网,通讯办公设备及用品,电子,电工,电器,机械,设备,仪器,仪表及专业用品,金属,非金属材料及制品,化工,交通,能源,矿产,冶金冶炼,粮油,食品,农林牧渔 
货运网，货运信息网，全国货运网，全国货运信息网，中国货运网，中国货运信息网，诚信，诚信物流，诚信物流信息，诚信物流网，诚信物流信息网，中国最大物流信息交易中心，物流信息客户端，物流信息中心，中国物流联盟网，物流联盟，物流联盟网,物流,中国物流,物流,中国物流,上海物流,企业物流,东莞物流,厦门物流,成都物流,广东物流,浙江物流,北京物流,江苏物流,广州物流,中华大物流,中国物流在线,中国企业物流,电话号簿,电信物流,工商物流,号簿,汕头物流,短信物流,114,在线物流,物流在线,网上物流,电子物流,光盘物流,移动物流,印刷物流,物流查询,物流搜索,网上114,企业物流,网上推广,网络营销,中小企业,行业目录,物流数据,全球物流,物流服务,物流指南,科学,教育,文化,医药,卫生,保健,公用事业,生活服务,餐饮,娱乐,购物,旅游,居家生活用品,服装,皮革,纺织,文体休闲用品,礼品,工艺品,商务服务,广告,物流,印刷,包装,造纸,工程,建筑,房地产,装潢,计算机,互联网,通讯办公设备及用品,电子,电工,电器,机械,设备,仪器,仪表及专业用品,电子商务,企业上网,供求信息,商业公告板,商业贸易机会,进出口贸易机会,贸易商情，商情广贴,商情广贴,翻译服务,信息咨询,域名注册,空间租用,网站建设,政府上网,中刮锪?工商信息,经济报道,行业报告,统计数据,企业信用,每日商机,贸易公告栏,企业电子商务,商业门户,企业网站策划，网站网页制作和维护，政府部门上网，经济技术开发区， 
配货网，配货信息网，全国配货网，全国配货信息网，中国配货网，中国配货信息网，货运，货运信息，全国货运，全国货运信息，中国货运，中国货运信息，>
<SCRIPT language=JavaScript type=text/JavaScript>
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</SCRIPT>
<META content="MSHTML 6.00.2462.0" name=GENERATOR>

</HEAD>
<BODY topMargin=0>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0 bgcolor="#FFFFFF">
  <TBODY> 
  <TR> 
    <TD vAlign=center align=middle width=203> 
      <div align="center"><a 
      href="http://www.cx56w.com"><img src="../images/logo.gif" border=0><br>
        </a><font color="#FF0000">今天是： 
        <script language=javascript>                                     
var day="";                                      
var month="";                                      
var ampm="";                                      
var ampmhour="";                                      
var myweekday="";                                      
var year="";                                      
mydate = new Date();                                      
myweekday = mydate.getDay();                                      
mymonth = mydate.getMonth();                                      
myday= mydate.getDate();                                      
myyear= mydate.getYear();                                      
myhours = mydate.getHours();                                      
ampmhour  =  (myhours > 12) ? myhours - 12 : myhours;                                      
if (ampmhour == "0") ampmhour = 0;                                      
ampm =  (myhours >= 12) ? ' 下午' : ' 上午';                                      
mytime = mydate.getMinutes();                                      
myminutes =  ((mytime < 10) ? ':0' : ':') + mytime;                                      
year = (myyear > 200) ? myyear : 1901 + myyear;                                      
if(myweekday == 0)                                      
weekday = " 星期日 ";                                      
else if(myweekday == 1)                                      
weekday = " 星期一 ";                                      
else if(myweekday == 2)                                      
weekday = " 星期二 ";                                      
else if(myweekday == 3)                                      
weekday = " 星期三 ";                                      
else if(myweekday == 4)                                      
weekday = " 星期四 ";                                      
else if(myweekday == 5)                                      
weekday = " 星期五 ";                                      
else if(myweekday == 6)       weekday = " 星期六 ";                                      
                            </script>
        <script>                                        
document.write(year + "年" + (mymonth+1) + "月" +myday + "日")            
document.write(" " +weekday + " ");</script>
        </font></div>
    </TD>
    <TD vAlign=top width=8></TD>
    <TD vAlign=top align=right>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td valign=center width="13%"> 
            <div align="right"><span 
            style="FONT-SIZE: 12px"></span><a 
            href="../Reg/User_Reg.asp" target=_blank><img 
            height=17 src="../images/register_button.gif" width=60 
            border=0></a></div>
          </td>
          <td valign=center align=middle width="12%"> 
            <div align="center"><span 
            style="FONT-SIZE: 12px"><font color=#000000><a 
            href="../Login/login.asp" target=_blank><img 
            height=17 src="../images/login_button.gif" width=40 
            border=0></a></font></span></div>
          </td>
          <td valign=center class=s height="30" width="75%" align="right"> 
            <div align="center"><a href="../index.asp">首页</a> ┆ <a href="../aboutus/index.asp">服务中心</a> 
              ┆ <a href="../agent/index.asp">加盟分站</a> ┆ <a href="/union/index.asp">各地分站</a> 
              ┆ <a onMouseOver="javascript:window.external.addFavorite('http://www.cx56w.com','诚信物流网')" href="http://www.zhwlw.org/">收藏本站</a> 
              ┆ <a onMouseOver="javascript:this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.cx56w.com');"
href="http://www.cx56w.com" target="_self">设为首页</a></div>

          </td>
        </tr>
      </table>
      <table cellspacing=0 cellpadding=0 border=0 width="100%">
        <tbody> 
        <tr align=right> 
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=92% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid" 
                align=center bgcolor=#f8f8f8><a 
                  href="../truck/index_beijing.asp" class=14link><strong><b>找货找车</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=92% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid" 
                align=center bgcolor=#f8f8f8><a 
                  href="../Member/goods_add.asp" class=14link ><strong><b>发布货源</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid" 
                align=center bgcolor=#f8f8f8><a class=14link 
            href="file:///D|/Member/car_add.asp"><strong><b>发布车源</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid"><a class=14link 
            href="../Repairs/index_beijing.asp"><strong><b>修理配件</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid"><a class=14link 
            href="../job/index_beijing.asp"><strong><b>招聘求职</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid"><a href="../yellowpage/index_beijing.asp"  class=14link><b>物流黄页</b></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td height=6 bgcolor=#ff6600></td>
              </tr>
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid"><a href="../news/index.asp" class=14link><b>物流资讯</b></a><a 
                  href="/info/"></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
      <table cellspacing=0 cellpadding=0 border=0 width="100%">
        <tbody> 
        <tr align=right> 
          <td height="30"> 
            <table height=24 cellspacing=0 cellpadding=0 width=92% border=0>
              <tbody> 
              <tr> 
                <td 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
                align=center bgcolor=#f8f8f8><a 
                  href="../fline/index_beijing.asp" class=14link><strong><b>货运专线</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=92% border=0>
              <tbody> 
              <tr> 
                <td 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"
                align=center bgcolor=#f8f8f8><a 
                  href="../pline/index_beijing.asp" class=14link ><strong><b>客运专线</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"
                align=center bgcolor=#f8f8f8><a class=14link 
            href="../cheliang/index_beijing.asp"><strong><b>车辆档案</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
               style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"><a class=14link 
            href="../info/index_beijing.asp"><strong><b>库房信息</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"><a class=14link 
            href="../trade/index_beijing.asp"><strong><b>供求信息</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"><a href="index.asp"  class=14link><b>网上地图</b></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"><a href="../member" class=14link><b>会员助手</b></a><a 
                  href="/info/"></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </TD>
  </TR>
  </TBODY>
</TABLE>
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
      <td align="center"><iframe id="ifrsearch" src="../search/index_car.asp" width="100%" frameborder="0" scrolling="no" height="90"></iframe></td>
    </tr>
  </table>
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<TABLE cellSpacing=0 width=778 align=center border=0>
  <TBODY>
  <TR>
      <TD vAlign=top width=290> 
        <table width="275" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td height="180" valign="top"> 
              <TABLE cellSpacing=1 cellPadding=0 width="100%" bgColor=#c4c2c2 
        border=0>
		        <TR>
          <TD vAlign=top align=middle bgColor=#ffffff>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD background=../images/login_bg.jpg height=29>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD class=00088 width="45%" height=21>
                        <DIV class=style10 align=center><font color="#FFFFFF">·会员登录</font></DIV></TD>
                      <TD width="55%"></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=4 width="100%" border=0>
              <TR>
                            <TD vAlign=top background=../images/login_bg1.gif 
                bgColor=#f6f6f6 height=144> 
                              <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
                   <FORM action=../login/Login_Check.asp method=POST name=loginForm> 
                    <TR>
                      <TD width="36%" height=31>
                        <DIV align=center>会员名：</DIV></TD>
                      <TD width="64%">
                                <INPUT id="username2" 
                        style="BORDER-RIGHT: #bbc0ca 1px solid; BORDER-TOP: #bbc0ca 1px solid; BORDER-LEFT: #bbc0ca 1px solid; BORDER-BOTTOM: #bbc0ca 1px solid; FONT-FAMILY: 宋体" 
                        size=13 name="UserID">
                              </TD></TR>
                    <TR>
                      <TD height=32>
                        <DIV align=center>密 码：</DIV></TD>
                      <TD>
					            <INPUT id="password" 
                        style="BORDER-RIGHT: #bbc0ca 1px solid; BORDER-TOP: #bbc0ca 1px solid; BORDER-LEFT: #bbc0ca 1px solid; BORDER-BOTTOM: #bbc0ca 1px solid; FONT-FAMILY: 宋体" 
                        type=password size=13 name="Pwd">
                                 </TD></TR><INPUT 
                    type=hidden value=login name=send> <INPUT type=hidden 
                    value=false name=B_code> 
                    <TR>
                      <TD colSpan=2 height=29>
                        <DIV align=center><INPUT class=noo type=image 
                        src="../images/bbslogin.gif">
                                  &nbsp; <a href="../Reg/User_Reg.asp"><IMG 
                        src="../images/bbsreg.gif" border=0></a></DIV>
                              </TD></TR>
                                  <TR align="center"> 
                                    <TD height=21 colSpan=2>发布货源信息│发布车辆信息</TD>
                                  </TR>
                    <TR>
                      <TD colSpan=2 height=26>
                        <DIV align=center>发布货源信息│发布车辆信息</DIV></TD></TR></FORM></TABLE>
                            </TD>
                          </TR></TABLE></TD></TR></TABLE>
			</td>
        </tr>
      </table>
        
      <table width=95% 
                    border=0 cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>最新客运专线</strong></font></td>
          <td width="130" height="22" valign="middle" align="center"> <a href="../Member/kyzx_add.asp"><font color="#FFFF00">免费发布客运专线信息</font></a></td>
          <td width="38" height="22" valign="middle" align="right"><a href="../product/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a></td>
        </tr>
        <tr> 
          <td colspan="3" valign="top" bgcolor="#FFFFFF" height="110"> 
            <%set rs_k=server.CreateObject("adodb.recordset")
		    rs_k.open "select top 10 * from zx_info where infotype='客运专线' and (city='北京' or city2='北京') order by infoid desc",conn,1,1
			
		  %>
            <table width="100%">
              <%if rs_k.bof and rs_k.eof then%>
              <tr> 
                <td height=23 align=left class=1>暂时无信息</td>
              </tr>
              <%else%>
              <%k=0%>
              <tr> 
			  <%do while not rs_k.eof%>
                <td height=23 align=left class=1><a href="../pline/detail.asp?InfoID=<%=rs_k("infoid")%>" target="_blank"><%=rs_k("city")%>←→<%=rs_k("city2")%></a></td>
              
              <%rs_k.movenext
			     k=k+1
				 if k mod 2=0 then
				 response.Write "</tr><tr>"
				 end if
			    loop
				rs_k.close
				set rs_k=nothing
			  %>
              <%end if%>
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
        
      <table width=95% 
                    border=0 cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>最新货运专线</strong></font></td>
          <td width="130" height="22" valign="middle" align="center"><a href="../Member/hyzx_add.asp"><font color="#FFFF00">免费发布货运专线信息</font></a> 
          </td>
          <td width="38" height="22" valign="middle" align="right"><font color="#FFFFFF"><a href="../product/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a></font></td>
        </tr>
        <tr> 
          <td colspan="3" valign="top" bgcolor="#FFFFFF" height="110"> 
            <%set rs_h=server.CreateObject("adodb.recordset")
					 rs_h.open "select top 10 * from zx_info where infotype='货运专线' and (city='北京' or city2='北京') order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='供应' Order by InfoID desc")
			  %>
            <table width="100%">
              <%if rs_h.bof and rs_h.eof then%>
              <tr> 
                <td height=23 align=center class=1>暂时无消息</td>
              </tr>
              <%else%>
              <%j=0%>
              <tr> 
			  <%do while not rs_h.eof%>
                <td height=23 align=left class=1><a href="../fline/detail.asp?infoid=<%=rs_h("infoid")%>" target="_blank"><%=rs_h("city")%>←→<%=rs_h("city2")%></a></td>
              <%rs_h.movenext
			   j=j+1
			   if j mod 2=0 then
			   response.Write "</tr><tr>"
			   end if
			    loop
				rs_h.close
				set rs_h=nothing
			  %>
              <%end if%>
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
        
      <table width=95% border=0 cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td height="20" valign="middle" bgcolor="#FF6600" width="150"><font 
            class=Ttitlew><strong><span class=Twhite14><img height=18 
            src="../images/yuanjian27.gif" 
            width=18></span></strong></font> <font color="#FFFFFF"><strong>最新加盟物流企业</strong></font></td>
          <td align="center" height="22" valign="middle"> <a href="../Reg/User_Reg.asp"><font color="#FFFF00">免费注册会员</font></a></td>
          <td width="36" height="22" valign="middle" align="center"><font color="#FFFFFF"><a href="../company/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a></font></td>
        </tr>
        <tr> 
          <td colspan="3" valign="top" bgcolor="#FFFFFF"> 
            <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 5 * from qyml where city='北京' order by id desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='供应' Order by InfoID desc")
				%>
            <%if rs_info.bof and rs_info.eof then%>
            <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0>
              <tr> 
                <td  valign=top height=67>无信息</td>
              </tr>
            </table>
            <%else %>
            <table cellspacing=2 cellpadding=2 border=0 width=100%>
              <%do while not rs_info.eof%>
              <tr align="center"> 
                <td><%=rs_info("qylb")%></td>
                <td><%=rs_info("city")%></td>
                <td><%=rs_info("qymc")%></td>
                <td><a href="../yellowpage/detail.asp?InfoID=<%=rs_info("ID")%>" target="_blank">详情</a></td>
              </tr>
              <%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>
            </table>
            <%end if%>
          </td>
        </tr>
        </tbody> 
      </table>
    </TD>
    <TD vAlign=top align="right"> 
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
          <tbody> 
          <tr> 
            <td class=white12 valign=center bgcolor=#ff6600 height=22> 
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="9%"> 
                    <div align=center><img height=20 
                        src="../images/huo.gif" width=21></div>
                  </td>
                  
                <td width="22%"><font color=#ffffff>最新货源信息</font></td>
                  
                <td width="69%"> <font color="#FFFF00">发布信息排行榜</font><font color="#FFFFFF"> 
                  | </font><a href="../Member/car_manage.asp"><font color="#FFFF00"></font></a> 
                  <a href="../Member/goods_add.asp"><font color="#FFFF00"> 免费发布货源信息</font></a><font color="#FFFFFF"> 
                  | </font><a href="../Member/goods_manage.asp"><font color="#FFFF00">管理货源信息</font></a></td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr align=middle> 
            <td 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                valign=top height=140> 
                <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 10 * from info2 where infotype='货源信息' and (city='北京' or city2='北京')order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='供应' Order by InfoID desc")
				%>	
				<%if rs_info.bof and rs_info.eof then%>			
              <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top height=67>无信息</td>
                </tr> 
              </table>
                <%else %>
<MARQUEE width="460" height="95" onmouseout=this.start() onmouseover=this.stop() direction="up" behavior="scroll" scrolldelay="500">		
            <table cellspacing=2 cellpadding=2 border=0 width=100%>
              <%do while not rs_info.eof%>
              <tr> 
                  
                <td  valign=top align="left"><%=rs_info("province")%>&nbsp;<%=rs_info("city")%>&nbsp;<%=rs_info("area")%></td>
				  
                <td  valign=top align="center">→</td>
				  
                <td  valign=top align="center"><%=rs_info("province2")%>&nbsp;<%=rs_info("city2")%>&nbsp;<%=rs_info("area2")%></td>
				  
                <td  valign=top align="center"><%=rs_info("cartype")%></td>
				  
                <td  valign=top align="center"><%=rs_info("carload")%>吨</td>
				  
                <td  valign=top align="center"><%=rs_info("addtime")%></td>
				  
                <td  align="center"><A href="../good/detail.asp?InfoID=<%=rs_info("InfoID")%>" target="_blank">详情</A></td>
                </tr> 
				<%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>				
              </table>
				 </MARQUEE>
			   <%end if%>			  
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="61%" height=24>
                  </td>
                  <td width="39%" height=24> 
                    <div align=center><a 
                        href="../truck/index.asp">更多&gt;&gt;</a></div>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          </tbody> 
        </table>
        <TABLE cellSpacing=0 width="100%" border=0>
          <TBODY> 
          <TR> 
            <TD vAlign=top width="48%" height=151> 
              <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TBODY> 
                <TR> 
                  <TD class=white12 vAlign=center bgColor=#ff6600 height=22> 
                    <TABLE cellSpacing=0 width="100%" border=0>
                      <TBODY> 
                      <TR> 
                        <TD width="9%"> 
                          <DIV align=center><IMG height=20 
                        src="../images/huo.gif" width=21></DIV>
                        </TD>
                        
                      <TD width="22%"><FONT color=#ffffff>最新空车信息</FONT></TD>
                        
                      <TD width="69%"> <font color="#FFFF00">发布信息排行榜</font><font color="#FFFFFF"> 
                        | </font><a href="../Member/car_manage.asp"><font color="#FFFF00"></font></a> 
                        <a href="../Member/car_add.asp"><font color="#FFFF00">免费发布空车信息</font></a><font color="#FFFFFF"> 
                        | </font><a href="../Member/car_manage.asp"><font color="#FFFF00">管理车源信息</font></a></TD>
                      </TR>
                      </TBODY> 
                    </TABLE>
                  </TD>
                </TR>
                <TR align=middle> 
                  <TD 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                vAlign=top height=140> 
                    <TABLE height=67 cellSpacing=0 cellPadding=0 width="100%" 
                  border=0>
                      <TBODY> 
                      <TR> 
                        <TD  vAlign=top height=67>
                <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 10 * from info2 where infotype='车源信息' and (city='北京' or city2='北京') order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='供应' Order by InfoID desc")
				%>	
				<%if rs_info.bof and rs_info.eof then%>			
              <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top height=67>无信息</td>
                </tr> 
              </table>
                <%else %>
<MARQUEE width="460" height="95" onmouseout=this.start() onmouseover=this.stop() direction="up" behavior="scroll" scrolldelay="500" class="p1">				
                        <table cellspacing=2 cellpadding=2 border=0 width=100%>
                          <%do while not rs_info.eof%>
                          <tr> 
                            <td  valign=top align="left"><%=rs_info("province")%><%=rs_info("city")%>&nbsp;<%=rs_info("area")%></td>
				            <td align="center">→</td>
				            <td align="center"><%=rs_info("province2")%><%=rs_info("city2")%>&nbsp;<%=rs_info("area2")%></td>
				            <td align="center"><%=rs_info("cartype")%></td>
				            <td align="center"><%=rs_info("carload")%>吨</td>
				            <td align="center"><%=rs_info("addtime")%></td>
				            <td align="center"><A href="../truck/detail.asp?InfoID=<%=rs_info("InfoID")%>" target="_blank">详情</A></td>
                </tr> 
				<%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>				
              </table>
				 </MARQUEE>
			   <%end if%>						
					    </TD>
                      </TR>
                      </TBODY> 
                    </TABLE>
                    <TABLE cellSpacing=0 width="100%" border=0>
                      <TBODY> 
                      <TR> 
                        <TD width="61%" height=24> 
                          <DIV align=center></DIV>
                        </TD>
                        <TD width="39%" height=24> 
                          <DIV align=center><A 
                        href="../good/index.asp">更多&gt;&gt;</A></DIV>
                        </TD>
                      </TR>
                      </TBODY> 
                    </TABLE>
                  </TD>
                </TR>
                </TBODY> 
              </TABLE>
            </TD>
          </TR>
          </TBODY> 
        </TABLE>
      <DIV align=center> 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
          <tbody> 
          <tr> 
            <td class=white12 valign=center bgcolor=#ff6600 height=22> 
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="9%"> 
                    <div align=center><img height=20 
                        src="../images/huo.gif" width=21></div>
                  </td>
                  <td width="40%"><font color=#ffffff>最新车辆档案</font></td>
                  <td width="51%"> 
                    <div align=right><font color="#FFFF00">免费登记车辆</font></div>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr align=middle> 
            <td 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                valign=top height=140> 
              <table height=134 cellspacing=0 cellpadding=0 width="100%" 
                  border=0>
                <tbody> 
                <tr>
                  <td  valign=top>
                    <table cellspacing=1 cellpadding=2 border=0 width=100% bgcolor="#CCCC66">
                      <tr bgcolor="#99CC33"> 
                        <td>车牌号</td>
                        <td align="center">车型</td>
                        <td  valign=top align="center">车型</td>
                        <td  valign=top align="center">车主</td>
                      </tr>
					  </table>				  </td>
                </tr>
                <tr> 
                  <td  valign=top height=100>
                <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 10 * from file_info a,qyml b where b.city='北京' and a.gsid=b.user order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='供应' Order by InfoID desc")
				%>	
				<%if rs_info.bof and rs_info.eof then%>			
              <table cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top>无信息</td>
                </tr> 
              </table>
                <%else %>
<MARQUEE width="460" height="120" onmouseout=this.start() onmouseover=this.stop() direction="up" behavior="scroll" scrolldelay="500" class="p1">				
                    <table cellspacing=1 cellpadding=2 border=0 width=100% bgcolor="#CCCC66">
                      <%do while not rs_info.eof%>
                      <tr bgcolor="#FFFFFF"> 
                        <td><%=rs_info("carnum")%></td>
                        <td align="center"><%=rs_info("carpp")%></td>
                        <td  valign=top align="center"><%=rs_info("cartype")%></td>
                        <td  valign=top align="center"><A href="../cheliang/detail.asp?InfoID=<%=rs_info("InfoID")%>" target="_blank">详情</A></td>
                      </tr>
                      <%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>
                    </table>
				 </MARQUEE>
			   <%end if%>				  </td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="61%" height=24> 
                    <div align=center></div>
                  </td>
                  <td width="39%" height=24> 
                    <div align=center><a 
                        href="../fline/index.asp">更多&gt;&gt;</a></div>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          </tbody> 
        </table>
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
          <tbody> 
          <tr> 
            <td class=white12 valign=center bgcolor=#ff6600 height=22> 
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="9%"> 
                    <div align=center><img height=20 
                        src="../images/huo.gif" width=21></div>
                  </td>
                  <td width="40%"><font color=#ffffff>最新会员信息</font></td>
                  <td width="51%"> 
                    <div align=right></div>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr align=middle> 
            <td 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                valign=top height=140> 
              <table height=67 cellspacing=0 cellpadding=0 width="100%" 
                  border=0>
                <tbody> 
                <tr> 
                  <td  valign=top height=67>
                <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 10 * from qyml where city='北京' order by id desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='供应' Order by InfoID desc")
				%>	
				<%if rs_info.bof and rs_info.eof then%>			
              <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top height=67>无信息</td>
                </tr> 
              </table>
                <%else %>
<MARQUEE width="460" height="95" onmouseout=this.start() onmouseover=this.stop() direction="up" behavior="scroll" scrolldelay="500" class="p1">				
                    <table cellspacing=2 cellpadding=2 border=0 width=100%>
                      <%do while not rs_info.eof%>
                      <tr align="center"> 
                        <td><%=rs_info("qylb")%></td>
				        <td><%=rs_info("city")%></td>
				        <td><%=rs_info("qymc")%></td>
				        <td><%=rs_info("name")%></td>
				        <td><%=rs_info("phone2")%>-<%=rs_info("phone3")%></td>
				        <td><A href="../yellowpage/detail.asp?InfoID=<%=rs_info("ID")%>" target="_blank">详情</A></td>
                </tr> 
				<%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>				
              </table>
				 </MARQUEE>
			   <%end if%>					  
				  </td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="61%" height=24> 
                    <div align=center></div>
                  </td>
                  <td width="39%" height=24> 
                    <div align=center><a 
                        href="../yellowpage/index.asp">更多&gt;&gt;</a></div>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          </tbody> 
        </table>
      </DIV>
    </TD>
  </TR></TBODY></TABLE>
<table cellspacing=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
      <td><img height=102 
      src="../images/banner2.gif" width=778 border=0></td>
  </tr>
  </tbody>
</table>
<table width="778" border="0" align="center" cellspacing="0">
  <tr> 
    <td width="545" valign="top">
      <table width=100% 
                    border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td width="118" height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>最新供求信息</strong></font></td>
          <td width="432" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF"><a href="../trade/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a></font></div>          
          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF"  height="200">
		  <%set rs_gq=server.CreateObject("adodb.recordset")
		  rs_gq.open "select top 8 * from gq_info where city='北京' order by infoid asc",conn,1,1
		  %>
		    <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <%if rs_gq.bof and rs_gq.eof then%>
              <tr>
              <td valign="top" colspan="4">暂时无信息</td>
            </tr>
            <%else%>
            <%do while not rs_gq.eof%>
            <tr>
              <td width="72" valign="top" align="center"><%=rs_gq("infotype")%></td>
              <td width="152" valign="top" align="center"><a href="../trade/detail.asp?infoid=<%=rs_gq("infoid")%>" target="_blank"><%=rs_gq("title")%></a></td>
              <td width="215" valign="top"><%=rs_gq("province")%>&nbsp;<%=rs_gq("city")%>&nbsp;<%=rs_gq("area")%></td>
              <td width="98" valign="top"><%=rs_gq("addtime")%></td>
            </tr>
            <%rs_gq.movenext
			    loop
				rs_gq.close
				set rs_gq=nothing
			  %>
            <%end if%>
          </table></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td width="217" valign="top">
      <table width=100% 
                    border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>招聘信息</strong></font></td>
          <td width="204" height="22" valign="middle">
            <div align="right"><font color="#FFFFFF"><a href="../company/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a></font></div>
          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF" height="200">
		  <%set rs_z=server.CreateObject("adodb.recordset")
		   rs_z.open "select top 8 * from zp_info where infotype='招聘信息' and city='北京' order by infoid desc",conn,1,1
		  %>
            <table border="0" cellspacing="0" cellpadding="2" width="100%">
              <%if rs_z.bof and rs_z.eof then%>
              <tr> 
                <td colspan="3" valign="top">暂时无信息</td>
              </tr>
			  <%else%>
			  <%do while not rs_z.eof%>
              <tr> 
                <td valign="top" align="center"> <a href="../job/detail.asp?infoid=<%=rs_z("infoid")%>" target="_blank"><%=rs_z("job")%></a></td>
                <td valign="top" align="center"><%=rs_z("province")%>&nbsp;<%=rs_z("city")%></td>
                <td valign="top" align="center"><%=rs_z("addtime")%></td>
              </tr>	
			  <%rs_z.movenext
			    loop
				rs_z.close
				set rs_z=nothing
			  %>
			  <%end if%>		  
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
</table>
<table width="778" border="0" align="center" cellspacing="0">
  <tr> 
    <td width="545" valign="top"> 
      <table width=100% 
                    border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>最新修理配件</strong></font></td>
          <td width="204" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF"><a href="../Repairs/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a></font></div>
          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF" height="200">
		  <%set file_info=server.CreateObject("adodb.recordset")
		    file_info.open "select top 8 * from repair_info where city='北京' order by infoid desc",conn,1,1
		  %>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <%if file_info.bof and file_info.eof then%>
              <tr> 
                <td colspan="3" valign="top">暂时无信息 </td>
              </tr>
			  <%else%>
              <%do while not file_info.eof%>
			  <tr align="center"> 
                <td width="107" valign="top" align="center">
				<a href="../repairs/detail.asp?infoid=<%=file_info("infoid")%>" target="_blank"><%=file_info("title")%></a> </td>
                <td width="172" valign="top"><%=file_info("province")%>&nbsp;<%=file_info("city")%>&nbsp;<%=file_info("area")%></td>
                <td width="256" valign="top"><%=file_info("addtime")%></td>
                </tr>	
			  <%file_info.movenext
			    loop
				file_info.close
				set file_info=nothing
			  %>
			  <%end if%>		  
            </table>		  
		  
		  </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td width="217" valign="top"> 
      <table width=100% 
                    border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>物流人才</strong></font></td>
          <td width="204" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF"><a href="../company/index.asp" target="_blank"><font color="#FFFFFF">更多</font></a></font></div>
          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF" height="200"> 
		  <%set rs_q=server.CreateObject("adodb.recordset")
		    rs_q.open "select top 8 * from zp_info where infotype='求职信息' and city='北京' order by infoid desc",conn,1,1
		  %>
            <table border="0" cellspacing="0" cellpadding="2" width="100%">
              <%if rs_q.bof and rs_q.eof then%>
              <tr> 
                <td colspan="3" valign="top">暂时无信息 </td>
              </tr>
			  <%else%>
              <%do while not rs_q.eof%>
			  <tr> 
                <td valign="top" align="center"> <a href="../job/detail.asp?infoid=<%=rs_q("infoid")%>" target="_blank"><%=rs_q("job")%></a> 
                </td>
                <td valign="top" align="center"><%=rs_q("province")%>&nbsp;<%=rs_q("city")%></td>
			    <td valign="top" align="center"><%=rs_q("addtime")%></td>
			  </tr>	
			  <%rs_q.movenext
			    loop
				rs_q.close
				set rs_q=nothing
			  %>
			  <%end if%>		  
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
</table>
<TABLE cellSpacing=0 width=778 align=center border=0>
  <TBODY>
  <TR>
    <TD><img height=113 
      src="../images/bannar.jpg" width=778 border=0></TD>
  </TR></TBODY></TABLE>
<table width="778" border="0" align="center" cellspacing="0">
  <tr> 
    <td width="554" valign="top"> 
      <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td width="48%" height="151" valign="top"> 
            <table cellspacing=0 cellpadding=0 width=100% border=0>
              <tbody> 
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        <div align="center"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">物流动态</font>　　　　　　　<a href="../user/login.asp" target="_blank"></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                <td height="127" valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
                  <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
          <tbody> 
          <%set rs_n=conn.execute("select top 8 * from NewsData where NClassID=1 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
            <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="../News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a> 
              (<%=month(rs_n("RegTime"))%>/<%=day(rs_n("RegTime"))%>)</td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
          </tbody> 
        </table>
        
      </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td width="1%" valign="top">&nbsp;</td>
          <td width="51%" valign="top"> 
            <table cellspacing=0 cellpadding=0 width=100% border=0>
              <tbody> 
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        <div align="center"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">物流知识</font>　　　　　　　<a href="../user/login.asp" target="_blank"></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                <td height="127" valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
                  <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
          <tbody> 
          <%set rs_n=conn.execute("select top 8 * from NewsData where NClassID=2 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
            <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="../News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a> 
              (<%=month(rs_n("RegTime"))%>/<%=day(rs_n("RegTime"))%>)</td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
          </tbody> 
        </table>
        
      </td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="220" valign="top"> 
      <table width=217 height="202" border=0 align="right" cellpadding=0 cellspacing=0>
        <tbody> 
        <tr> 
          <td height="22" bgcolor="#FF6600"><font face=黑体 size=3> 　<font color="#FFFFFF" size="2"><strong>物流工具</strong></font></font> 
          </td>
        </tr>
        <tr> 
          <td height="184" valign="top" 
          style="BORDER-RIGHT: #cccccc 1px solid; PADDING-RIGHT: 8px; BORDER-TOP: #cccccc 3px solid; PADDING-LEFT: 8px; FONT-SIZE: 12px; PADDING-BOTTOM: 10px; BORDER-LEFT: #cccccc 1px solid; LINE-HEIGHT: 150%; PADDING-TOP: 10px; BORDER-BOTTOM: #cccccc 1px solid"> 
            <table width="100%" height="80" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hntrans.com:81/gs/xinlichsele.php" target="_blank" class="blacklink">里程查询</a>]</td>
                <td height="22" align="center">　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.mapbar.com:8000/commonmapright/?title=浙江物流网&cityCode=0571&width=760&zm=20&logo=true&cityOption=ture&mapstyle=0&state=map666&advscrolling=false&color=imgblue&headType=lydt&headPic=false" target="_blank" class="blacklink">物流地图</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://cb.kingsoft.com/" target="_blank" class="blacklink">英汉工具</a>]</td>
                <td height="22" align="center">　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://zj.183.com.cn/quire/code.htm" target="_blank" class="blacklink">邮编查询</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hlbr-l-tax.gov.cn/xxfw/slcx/001%5B1%5D.htm" target="_blank" class="blacklink">税率查询]</a></td>
                <td height="22" align="center">　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://info./jinchu/whhx/default.shtml" target="_blank">结汇核销</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hzti.com/" target="_blank" class="blacklink">交通状况</a>]</td>
                <td height="22" align="center"> 　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.xichang.tv/calendar.htm" target="_blank" class="blacklink">电子日历</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.5l.com.cn/second/shouce/kgdm.htm" target="_blank">国际空港</a>]</td>
                <td height="22" align="center">　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.5l.com.cn/second/shouce/gonglushoufei.htm" target="_blank" class="blacklink">公路收费</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="../tool/sjhb.htm" target="_blank">世界货币</a>]</td>
                <td height="22" align="center"> 　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://info./entr/list_cgs.shtml" target="_blank">货物追踪</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="../tg/bgdm.asp" target="_blank">报关代码</a>]</td>
                <td height="22" align="center">　<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.chemaid.com/sample/user.asp" target="_blank" class="blacklink">危险品名</a>]</td>
              </tr>
            </table>
                </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
</table>
<table width="778" border="0" align="center" cellspacing="0">
  <tr> 
    <td width="554" valign="top"> 
      <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td width="48%" height="151" valign="top"> 
            <table cellspacing=0 cellpadding=0 width=100% border=0>
              <tbody> 
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        <div align="center"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">企业播报</font>　　　　　　　<a href="../user/login.asp" target="_blank"></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                <td height="127" valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
                  <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
          <tbody> 
          <%set rs_n=conn.execute("select top 8 * from NewsData where NClassID=3 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
            <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="../News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a> 
              (<%=month(rs_n("RegTime"))%>/<%=day(rs_n("RegTime"))%>)</td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
          </tbody> 
        </table>
      </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td width="1%" valign="top">&nbsp;</td>
          <td width="51%" valign="top"> 
            <table cellspacing=0 cellpadding=0 width=100% border=0>
              <tbody> 
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        <div align="center"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">物流会展</font>　　　　　　　<a href="../user/login.asp" target="_blank"></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                <td height="127" valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
                  <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
          <tbody> 
          <%set rs_n=conn.execute("select top 8 * from NewsData where NClassID=4 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
            <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="../News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a> 
              (<%=month(rs_n("RegTime"))%>/<%=day(rs_n("RegTime"))%>)</td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
          </tbody> 
        </table>
      </td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="220" valign="top"> 
      <table width=217 height="202" border=0 align="right" cellpadding=0 cellspacing=0>
        <tbody> 
        <tr> 
          <td height="22" bgcolor="#FF6600"><font face=黑体 size=3> 　<font color="#FFFFFF" size="2"><strong>便民工具</strong></font></font> 
          </td>
        </tr>
        <tr> 
          <td height="180" valign="top" 
          style="BORDER-RIGHT: #cccccc 1px solid; PADDING-RIGHT: 8px; BORDER-TOP: #cccccc 3px solid; PADDING-LEFT: 8px; FONT-SIZE: 12px; PADDING-BOTTOM: 10px; BORDER-LEFT: #cccccc 1px solid; LINE-HEIGHT: 150%; PADDING-TOP: 10px; BORDER-BOTTOM: #cccccc 1px solid"> 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
              <tbody> 
              <tr align="center"> 
                <td height=20><a href="http://www.t7online.com/China.htm" 
      target=_blank>[天气预报]</a></td>
                <td height="20"><a href="http://www.21page.net/public/code_country.asp" target="_blank">[邮编区号</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a 
                        href="http://www.go2map.com" target="_blank">中国地图</a>]</td>
                <td height="20"><a href="../tool/rili.htm" target="_blank">[万 年 历]</a></td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a 
                        href="http://www.id5.cn/uindex.jsp?webid=cx56w" target="_blank">身份核查</a>]</td>
                <td height="20">[<a href="http://www.hntrans.com:81/gs/xinlichsele.php" target="_blank">里程查询</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a href="http://train.chinamor.cn.net/cctt.asp" target="_blank">列车时刻</a>]</td>
                <td height="20">[<a 
                        href="http://www.123cha.com" target="_blank">手机查询</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a href="http://www.21page.net/public/time_world.asp" target="_blank">世界时间</a>]</td>
                <td height="20">[<a href="../tool/road.htm" 
      target=_blank>国道查询</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a href="../tool/meipay.asp" target="_blank"><font color="#FF0000">网上银行</font></a>]</td>
                <td height="20">[<a class=blue 
                  href="http://ent.tom.com/tv/forecast/2.html" 
                  target=_blank>电视节目</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a class=blue 
                  href="http://act.it.sohu.com/download/download.php" 
                  target=_blank>装机软件</a>]</td>
                <td height="20">[<a class=blue 
                  href="http://www.chinaren.com/s2005/topmusic.shtml" 
                  target=_blank>音乐排行</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a 
                        href="http://www.alexa.com/site/ds/top_sites?ts_mode=lang&lang=zh_gb2312" target="_blank" class=titlefont>排名查询</a>]</td>
                <td height="20">[<font color="#FFFFFF"><font color="#FFFFFF"><font color="#FFFFFF"><a href="http://site.cx56w.com" target="_blank"><font color="#FF0000">更多便民</font></a></font></font></font>]</td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
</table>
<table border="0" width="778" id="table2" height="80" cellpadding="2" align="center" cellspacing="1" bgcolor="#ff6600">
  <tr> 
    <th bgcolor="#FF9900" width="22" align="center"> 
      <p align="right"><font size="3" color="#FFFFFF">各地分站</font></p>
    </th>
    <td bgcolor="#FFFFFF"> 
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FF9900">
        <!--DWLayoutTable-->
        <tr align="center" valign="middle"> 
    <td width="69" height="23" bgcolor="#FFFFFF"><a href="beijing.asp" target="_blank">北京分网</a></td>
    <td width="69" bgcolor="#FFFFFF"><a href="shanghai.asp" target="_blank">上海分网</a></td>
    <td width="69" bgcolor="#FFFFFF"><a href="tianjin.asp" target="_blank">天津分网</a></td>
    <td width="69" bgcolor="#FFFFFF"><a href="chongqing.asp" target="_blank">重庆分网</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="shangdong.asp" target="_blank">山东分网</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="shanxi.asp" target="_blank">山西分网</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="henan.asp" target="_blank">河南分网</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="hebei.asp" target="_blank">河北分网</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="hunan.asp" target="_blank">湖南分网</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="hubei.asp" target="_blank">湖北分网</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="hainan.asp" target="_blank">海南分网</a></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td height="23" align="center" valign="middle" bgcolor="#FFFFFF"><a href="jiangxi.asp" target="_blank">江西分网</a></td>
    <td bgcolor="#FFFFFF"><a href="sichuan.asp" target="_blank">四川分网</a></td>
    <td bgcolor="#FFFFFF"><a href="anhui.asp" target="_blank">安徽分网</a></td>
    <td bgcolor="#FFFFFF"><a href="fujian.asp" target="_blank">福建分网</a></td>
    <td bgcolor="#FFFFFF"><a href="jiangsu.asp" target="_blank">江苏分网</a></td>
    <td bgcolor="#FFFFFF"><a href="guangxi.asp">广西分网</a></td>
    <td bgcolor="#FFFFFF"><a href="guangdong.asp">广东分网</a></td>
    <td bgcolor="#FFFFFF"><a href="guizhou.asp" target="_blank">贵州分网</a></td>
    <td bgcolor="#FFFFFF"><a href="yunnan.asp" target="_blank">云南分网</a></td>
    <td bgcolor="#FFFFFF"><a href="zhejiang.asp" target="_blank">浙江分网</a></td>
    <td bgcolor="#FFFFFF"><a href="shan xi.asp" target="_blank">陕西分网</a></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td height="25" bgcolor="#FFFFFF"><a href="gansu.asp" target="_blank">甘肃分网</a></td>
    <td bgcolor="#FFFFFF"><a href="qinghai.asp" target="_blank">青海分网</a></td>
    <td bgcolor="#FFFFFF"><a href="liaoning.asp" target="_blank">辽宁分网</a></td>
    <td bgcolor="#FFFFFF"><a href="heilongjiang.asp" target="_blank">黑龙江网</a></td>
    <td bgcolor="#FFFFFF"><a href="jilin.asp" target="_blank">吉林分网</a></td>
    <td bgcolor="#FFFFFF"><a href="ningxia.asp" target="_blank">宁夏分网</a></td>
    <td bgcolor="#FFFFFF"><a href="neimenggu.asp" target="_blank">内蒙分网</a></td>
    <td bgcolor="#FFFFFF"><a href="xinjiang.asp" target="_blank">新疆分网</a></td>
    <td bgcolor="#FFFFFF"><a href="xizang.asp" target="_blank">西藏分网</a></td>
    <td bgcolor="#FFFFFF"><a href="../index.asp" target="_blank">物流联盟</a></td>
    <td bgcolor="#FFFFFF"><a href="http://www.cx56w.com">联盟总站</a></td>
  </tr>
</table>
    </td>
  </tr>
</table>
<TABLE width="778" border="0" cellspacing="0" cellpadding="2" align="center">
  <TBODY> 
  <TR> 
    <TD align=center height="10"></TD>
  </TR>
  <TR> 
    <TD align=center height="3" bgcolor="#FF6600"></TD>
  </TR>
  <TR> 
    <TD height="26" align=center bgcolor="#F7F7F7"><a 
      href="../aboutus/index.asp">关于我们</a> | <a 
      href="../aboutus/service.asp" target=_blank>服务条款</a> | <a 
      href="../aboutus/guide.asp">网站指南</a> | <a 
      href="../aboutus/hyzh.asp">会员资费</a> | <a 
      href="../aboutus/exp.asp">支付方式</a> |　<a 
      href="../aboutus/ad.asp">网络广告</a> | <a 
      href="../aboutus/cooperate.asp">市场合作</a> | <a 
      href="../aboutus/supply.asp">代理商专区</a></TD>
  </TR>
  </TBODY> 
</TABLE>
<table cellspacing=2 cellpadding=2 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td valign=center align=center>CopyRight &copy; 2003-2006 <a href="http://www.cx56w.com" target="_blank">Www.Cx56w.Com</a> 
      All Rights Reserved.<nobr></nobr> </td>
  </tr>
  <tr> 
    <td valign=center align=center>服务热线：13116601816 电子邮箱：info@cx56w.com</td>
  </tr>
  <tr> 
    <td valign=center align=center>网络实名：诚信物流网 备案号：浙ICP备05059371号</td>
  </tr>
  </tbody> 
</table>
</BODY></HTML>
