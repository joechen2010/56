<!--#include file="Inc/Conn2.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
const MaxNumber=50
const MaxNum=8
const MaxLen=30
dim rs_sel,sel1_total,rs_buy,buy_total,rs_p,p_total
set rs_sel=conn.execute("select Count(*) from Info where InfoType='��Ӧ'")
sell_total=rs_sel(0)
rs_sel.close
set rs_sel=nothing

set rs_buy=conn.execute("select Count(*) from Info where InfoType='�ɹ�'")
buy_total=rs_buy(0)
rs_buy.close
set rs_buy=nothing

set rs_s_o=conn.execute("select Count(*) from offer where OfferType='��Ӧ'")
sell_o_total=rs_s_o(0)
rs_s_o.close
set rs_s_o=nothing

set rs_b_o=conn.execute("select Count(*) from offer where OfferType='�ɹ�'")
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
<TITLE>168������Ϣ��--ȫ����רҵ����ŵ�������Ϣ����ƽ̨|����,������,�й�������,������Ϣ��</TITLE>
<meta name=keywords content="����,������,�й�������,168������Ϣ��">
<meta name=description content="����-168������Ϣ���Ǵ��͹���������Ѷ��������ƷӦ��,��Ʒ���ҽ��ܵ�����ý��.��ӭ��Ϊ168������Ϣ��ע���Ա�ṩ�������!Find the logistics products and services in this website!">
<LINK href="images/css.css" type=text/css rel=stylesheet>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content=. name=description>
<META  content=��>
<SCRIPT language=JavaScript type=text/JavaScript>
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
function openwindow(url,width,height) { 	left1 =0; 	top1 = 0; 	window.open(url,"","width=" + width + ",height=" + height + ",toolbar=no,menubar=no,scrollbars=yes,location=no,status=yes,left=" + left1.toString() + ",top=" + top1.toString()); }
//resizable=yes, ���
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
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
<script language="javascript" src="images/ad.js"></script>
<SCRIPT FOR="sinadl" EVENT="FSCommand()" LANGUAGE="JavaScript">
	adRightFloat.style.visibility='hidden';
	adLeftFloat.style.visibility='hidden';
</SCRIPT>
<!--#include file="top.asp"-->
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<table cellspacing=0 cellpadding=0 align=center border=0 width=778>
  <tr align="center"> 
    <td height="66" background=images/bg_ss05_060516.gif><a href="#"><img src="images/G06.gif" width="138" height="60" border="1"></a></td>
    <td height="66" background=images/bg_ss05_060516.gif><a href="http://www.chinawhwl.com/zongshu/zhanhuijieshao.asp" target="_blank"><img src="images/G07.gif" width="138" height="60" border="1"></a></td>
    <td height="66" background=images/bg_ss05_060516.gif><a href="http://www.bjstb.com/" target="_blank"><img src="images/G03.gif" width="138" height="60" border="1"></a></td>
    <td height="66" background=images/bg_ss05_060516.gif><img src="images/shanbei.gif" width="148" height="60" border="1"></td>
    <td height="66" background=images/bg_ss05_060516.gif><img src="images/huaan.gif" width="148" height="60" border="1"></td>
  </tr>
</table>
<tr> 
  <td align="center"> 
    <TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>
      <TR> 
        <TD colspan="5" background=images/bg_ss05_060516.gif> 
          <table id=Table1 cellspacing=0 cellpadding=0 width="100%" align=center 
      border=0>
            <form method="POST" action="index.asp" name="form1">
              <tr> 
                <td height="30" align="right" bgcolor="#FF9900"><span class="STYLE1">������:</span> 
                  <select id="s1" name="province">
                    <option value="ʡ��">ʡ��</option>
                  </select>
                  <select id="s2" name="city">
                    <option>�ؼ���</option>
                  </select>
                  <select id="s3" name="area">
                    <option>�С��ؼ��С���</option>
                  </select>
                  &nbsp; 
                  <select name="type">
                    <option value="��Դ" <%if typeselect="��Դ" then response.Write "selected" end if%>>��Դ</option>
                    <option value="�ճ�" <%if typeselect="�ճ�" then response.Write "selected" end if%>>�ճ�</option>
                    <option value="�����" <%if typeselect="�����" then response.Write "selected" end if%>>�����</option>
                    <option value="����" <%if typeselect="����" then response.Write "selected" end if%>>����</option>
                    <option value="����" <%if typeselect="����" then response.Write "selected" end if%>>����</option>
                    <option value="��ҵ" <%if typeselect="��ҵ" then response.Write "selected" end if%>>��ҵ</option>
                  </select>
                  &nbsp; 
                  <input type="submit" value="�� ��"  name="button">
                  <input type="hidden" name="action" value="rolling">
                </td>
              </tr>
              <SCRIPT language=javascript>
            setup();
          </SCRIPT>
            </form>
          </table>
        </td>
      </tr>
    </table>
    <table width="92%" border="0" height="5" cellpadding="0" cellspacing="0">
      <tr> 
        <td></td>
      </tr>
    </table>
    <TABLE cellSpacing=0 width=778 align=center border=0>
      <TR> 
        <TD vAlign=top width=280> 
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
                          <TD background=images/login_bg.jpg height=29> 
                            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                              <TBODY> 
                              <TR> 
                                <TD class=00088 width="45%" height=21> 
                                  <DIV class=style10 align=center><font color="#FFFFFF">����Ա��¼</font></DIV>
                                </TD>
                                <TD width="55%"></TD>
                              </TR>
                              </TBODY> 
                            </TABLE>
                          </TD>
                        </TR>
                        </TBODY> 
                      </TABLE>
                      <%if session("user")="" then%>
                      <TABLE cellSpacing=0 cellPadding=4 width="100%" border=0>
                        <TR> 
                          <TD vAlign=top background=images/login_bg1.gif bgColor=#f6f6f6 height=144> 
                            <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center  border=0>
                              <FORM action=login/Login_Check.asp method=POST name=loginForm>
                                <TR> 
                                  <TD width="36%" height=31> 
                                    <DIV align=center>��Ա����</DIV>
                                  </TD>
                                  <TD width="64%"> 
                                    <INPUT id="username2" 
                        style="BORDER-RIGHT: #bbc0ca 1px solid; BORDER-TOP: #bbc0ca 1px solid; BORDER-LEFT: #bbc0ca 1px solid; BORDER-BOTTOM: #bbc0ca 1px solid; FONT-FAMILY: ����" 
                        size=13 name="UserID">
                                  </TD>
                                </TR>
                                <TR> 
                                  <TD height=32> 
                                    <DIV align=center>�� �룺</DIV>
                                  </TD>
                                  <TD> 
                                    <INPUT id="password" 
                        style="BORDER-RIGHT: #bbc0ca 1px solid; BORDER-TOP: #bbc0ca 1px solid; BORDER-LEFT: #bbc0ca 1px solid; BORDER-BOTTOM: #bbc0ca 1px solid; FONT-FAMILY: ����" 
                        type=password size=13 name="Pwd">
                                  </TD>
                                </TR>
                                <INPUT 
                    type=hidden value=login name=send>
                                <INPUT type=hidden 
                    value=false name=B_code>
                                <TR> 
                                  <TD colSpan=2 height=29> 
                                    <DIV align=center> 
                                      <INPUT class=noo type=image 
                        src="images/bbslogin.gif">
                                      &nbsp; <a href="Reg/User_Reg.asp"><IMG 
                        src="images/bbsreg.gif" border=0></a></DIV>
                                  </TD>
                                </TR>
                                <TR align="center"> 
                                  <TD height=21 colSpan=2>������Դ��Ϣ������������Ϣ</TD>
                                </TR>
                                <TR> 
                                  <TD colSpan=2 height=26> 
                                    <DIV align=center>������Դ��Ϣ������������Ϣ</DIV>
                                  </TD>
                                </TR>
                              </FORM>
                            </TABLE>
                          </TD>
                        </TR>
                      </TABLE>
                      <%else%>
                      <TABLE cellSpacing=0 cellPadding=4 width="100%" border=0>
                        <TR> 
                          <TD vAlign=top background=images/login_bg1.gif bgColor=#f6f6f6 height=144> 
                            <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
                              <TR align="center"> 
                                <TD width="100%" height=21>��ã�<%=session("name")%></TD>
                              </TR>
                              <TR> 
                                <TD height=26 align="center"><a href="site/aboutus.asp?InfoID=<%=session("id")%>&user=<%=session("user")%>" target="_blank">�ҵ���վ</a></TD>
                              </TR>
                              <TR> 
                                <TD height=26 align="center"><a href="Member/index.asp">��Ա����</a></TD>
                              </TR>
                              <TR> 
                                <TD height=26 align="center"> <a href="javascript:openwindow:openwindow('infocenter/infocenter_index.asp',1014,680)">��Ϣ����</a></TD>
                              </TR>
                              <TR> 
                                <TD height=26 align="center"><a href="Login/Logout.asp">ע����½</a></TD>
                              </TR>
                            </TABLE>
                          </TD>
                        </TR>
                      </TABLE>
                      <%end if%>
                    </TD>
                  </TR>
                </TABLE>
              </td>
            </tr>
          </table>
          <table width="92%" border="0" height="5" cellpadding="0" cellspacing="0">
            <tr> 
              <td></td>
            </tr>
          </table>
          <table width=275 border=0 cellpadding=0 cellspacing=0>
            <tr> 
              <td height="27" valign="baseline" colspan="2" background="images/bj01.gif"> 
                <table width="100%" border="0" cellspacing="2" cellpadding="0" height="27">
                  <tr> 
                    <td width="30" align="center"><img height=18 
            src="images/yuanjian27.gif" 
            width=18></td>
                    <td><font color="#FFFFFF"><strong>��վ����</strong></font></td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <table width=275 border=0 cellpadding=0 cellspacing=1 bgcolor="#FF6600">
            <tbody> 
            <tr> 
              <td colspan="3" valign="top" bgcolor="#FFFFFF" height="110"> 
                <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
                  <tbody> 
                  <%set rs_n=conn.execute("select top 5 * from NewsData where NClassID=7 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
                  <tr> 
                    <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No 
                      Data!</td>
                  </tr>
                  <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
                  <tr> 
                    <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="javascript:openwindow('News/NewsDetail1.asp?NewsID=<%=rs_n("NewsID")%>',500,400)"><%=topic%></a> 
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
          <table width="92%" border="0" height="5" cellpadding="0" cellspacing="0">
            <tr> 
              <td></td>
            </tr>
          </table>
          <table width=275 border=0 cellpadding=0 cellspacing=0>
            <tr> 
              <td height="27" valign="baseline" colspan="2" background="images/bj01.gif"> 
                <table width="100%" border="0" cellspacing="2" cellpadding="0" height="27">
                  <tr> 
                    <td><img height=18 
            src="images/yuanjian27.gif" 
            width=18></td>
                    <td><font color="#FFFFFF"><strong>�Ƽ�����ר��</strong></font></td>
                    <td><a href="Member/hyzx_add.asp"><font color="#FFFF00">��ѷ�������ר����Ϣ</font></a></td>
                    <td><font color="#FFFFFF"><a href="product/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></font></td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <table width=275 border=0 cellpadding=0 cellspacing=1 bgcolor="#999999">
            <tbody> 
            <tr> 
              <td colspan="3" valign="top" bgcolor="#FFFFFF"> 
                <%set rs_h=server.CreateObject("adodb.recordset")
					 rs_h.open "select top 10 * from zx_info where infotype='����ר��' order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='��Ӧ' Order by InfoID desc")
			  %>
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <%if rs_h.bof and rs_h.eof then%>
                  <tr> 
                    <td height=23 align=center class=1>��ʱ����Ϣ</td>
                  </tr>
                  <%else%>
                  <%j=0%>
                  <tr> 
                    <%do while not rs_h.eof%>
                    <td align=center class=1> <A href="javascript:openwindow('fline/detail.asp?infoid=<%=rs_h("infoid")%>',500,400)"> 
                      <%=rs_h("city")%>����<%=rs_h("city2")%></a></td>
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
          <table width="92%" border="0" height="5" cellpadding="0" cellspacing="0">
            <tr> 
              <td></td>
            </tr>
          </table>
          <table width=275 border=0 cellpadding=0 cellspacing=0>
            <tr> 
              <td height="27" colspan="2" background="images/bj01.gif"> 
                <table width=275 border=0 cellpadding=0 cellspacing=2>
                  <tbody> 
                  <tr bgcolor="#FF6600"> 
                    <td bgcolor="#FF6600" width="30"><img height=18 
            src="images/yuanjian27.gif" 
            width=18> <font color="#FFFFFF"></font></td>
                    <td bgcolor="#FF6600"><font color="#FFFFFF"><strong>�Ƽ�������ҵ</strong></font></td>
                    <td align="center" valign="middle"> <a href="Reg/User_Reg.asp"><font color="#FFFF00">���ע���Ա</font></a></td>
                    <td width="36" valign="middle" align="center"><font color="#FFFFFF"><a href="company/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></font></td>
                  </tr>
                  </tbody> 
                </table>
              </td>
            </tr>
          </table>
          <table width=275 border=0 cellpadding=0 cellspacing=1 bgcolor="#999999">
            <tbody> 
            <tr> 
              <td colspan="3" valign="top" bgcolor="#FFFFFF"> 
                <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 10 * from qyml order by id desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='��Ӧ' Order by InfoID desc")
				%>
                <%if rs_info.bof and rs_info.eof then%>
                <table cellspacing=0 cellpadding=0 width="100%" border=0>
                  <tr> 
                    <td  valign=top>����Ϣ</td>
                  </tr>
                </table>
                <%else %>
                <table cellspacing=0 cellpadding=0 border=0 width=100%>
                  <% h=0%>
                  <tr> 
                    <%do while not rs_info.eof%>
                    <td height="20"><font color="#006600">��<%=rs_info("city")%>��</font><%if len(rs_info("qymc"))>6 then%><a href="site/aboutus.asp?InfoID=<%=rs_info("ID")%>&user=<%=rs_info("user")%>" target="_blank"><%=left(rs_info("qymc"),6)%></a> 
                      <%else%><a href="site/aboutus.asp?InfoID=<%=rs_info("ID")%>&user=<%=rs_info("user")%>" target="_blank"><%=rs_info("qymc")%></a> 
                      <%end if%>
                    </td>
                    <%h=h+1
			  if h mod 2=0 then response.Write "</tr><tr>" end if
			  rs_info.movenext
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
          <table width="92%" border="0" height="5" cellpadding="0" cellspacing="0">
            <tr> 
              <td></td>
            </tr>
          </table>
          <%'*********************�����վ***********************%>
          <table width=275 border=0 cellpadding=0 cellspacing=0>
            <tr> 
              <td height="27" valign="baseline" colspan="2" background="images/bj01.gif"> 
                <table width="100%" border="0" cellspacing="2" cellpadding="0" height="27">
                  <tr> 
                    <td><img height=18 
            src="images/yuanjian27.gif" 
            width=18></td>
                    <td><font color="#FFFFFF"><strong>�����վ</strong></font></td>
                    <td><font color="#FFFF00">���˷�վ����</font></td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <table width=275 border=0 cellpadding=0 cellspacing=1 bgcolor="#999999">
            <tr> 
              <td colspan="2" valign="top" bgcolor="#FFFFFF"> 
                <table cellspacing=0 cellpadding=0 width="100%" border=0>
                  <tr> 
                    <td  valign=top>������վ</td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
        </TD>
        <%'//////////////////////////����///////////////////////////////%>
        <TD vAlign=top align="right"> 
          <%if typeselect="" then%>
          <!--#include file="middle.asp"-->
          <%end if%>
          <%if typeselect="��Դ" then%>
          <!--#include file="search_goods.asp"-->
          <%end if%>
          <%if typeselect="�ճ�" then%>
          <!--#include file="search_car.asp"-->
          <%end if%>
          <%if typeselect="�����" or typeselect="����" or typeselect="����" or typeselect="��ҵ" then%>
          <!--#include file="search_login.asp"-->
          <%end if%>
          <table cellspacing=0 width="100%" border=0>
            <tr> 
              <td width="24%" height=24 align="center"></td>
              <td width="76%" align="center"> 
                <%if request("action")="" then%>
                <a href="?action=rolling">��������Ϣ������ҳ��</a> 
                <%else %>
                <a href="?">���ص�������Ϣ����ҳ��</a> 
                <%end if%>
              </td>
            </tr>
          </table>
        </TD>
        <%'//////////////////////////////////////////////////////////////////%>
      </TR>
    </TABLE>
    <table width="778" border="0" cellspacing="0" align="center">
      <tr> 
        <td> 
          <div align="left"> <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="778" height="70">
              <param name="movie" value="truck/images/top6.swf">
              <param name="quality" value="high">
              <embed src="truck/images/top6.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="778" height="70">
              </embed> 
            </object> </div>
        </td>
      </tr>
    </table>
    <table width="778" border="0" align="center" cellspacing="0">
      <tr> 
        <td valign="top"> 
          <!--#include file="file_gq.asp"-->
        </td>
      </tr>
    </table>
    <table width="778" border="0" align="center" cellspacing="0">
      <tr> 
        <td valign="top"> 
          <!--#include file="job.asp"-->
        </td>
      </tr>
    </table>
    <table width="778" border="0" cellspacing="0" align="center">
      <tr> 
        <td> 
          <div align="left"> <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="778" height="70">
              <param name="movie" value="images/cx56w.swf">
              <param name="quality" value="high">
              <embed src="images/cx56w.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="778" height="70">
              </embed> 
            </object> </div>
        </td>
      </tr>
    </table>
    <table width="778" border="0" align="center" cellspacing="0">
      <tr> 
        <td valign="top"> 
          <table width="100%" border="0" cellspacing="0">
            <tr> 
              <td height="151" valign="top"> 
                <!--#include file="other.asp"-->
              </td>
            </tr>
          </table>
          <table width=217 height="202" border=0 align="right" cellpadding=0 cellspacing=0>
          </table>
        </td>
      </tr>
    </table>
    <!--#include file="bottom.html"-->
    <!-- Live800���߿ͷ�ͼ��:cx56w1[������] ��ʼ-->
    <script language="javascript" src="http://chat.live800.com:8080/live800/chatClient/floatButton.js?jid=9756223015&companyID=32913&configID=24234&codeType=custom"></script>
    <!-- Live800���߿ͷ�ͼ��:cx56w1[������] ����-->
    <!-- Live800���߿ͷ����ܴ��� ��ʼ-->
    <script language="javascript" src="http://chat.live800.com:8080/live800/chatClient/monitor.js?jid=9756223015&companyID=32913&configID=23583&codeType=custom"></script>
    <!-- Live800���߿ͷ����ܴ��� ����-->
</BODY>
</HTML>
