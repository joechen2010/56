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
session("fz")=request("fz")
fz=session("fz")
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
<script type="text/javascript">
var delta=0.15;
var collection;
var closeB=false;
 function floaters() {
  this.items = [];
  this.addItem = function(id,x,y,content)
      {
     document.write('<DIV id='+id+' style="Z-INDEX: 10; POSITION: absolute;  width:80px; height:80px;left:'+(typeof(x)=='string'?eval(x):x)+';top:'+(typeof(y)=='string'?eval(y):y)+'">'+content+'</DIV>');
     
     var newItem    = {};
     newItem.object   = document.getElementById(id);
     newItem.x    = x;
     newItem.y    = y;

     this.items[this.items.length]  = newItem;
      }
  this.play = function()
      {
     collection    = this.items
     setInterval('play()',1);
      }
  }
  function play()
  {
   if(screen.width<=800 || closeB)
   {
    for(var i=0;i<collection.length;i++)
    {
     collection[i].object.style.display = 'none';
    }
    return;
   }
   for(var i=0;i<collection.length;i++)
   {
    var followObj  = collection[i].object;
    var followObj_x  = (typeof(collection[i].x)=='string'?eval(collection[i].x):collection[i].x);
    var followObj_y  = (typeof(collection[i].y)=='string'?eval(collection[i].y):collection[i].y);

    if(followObj.offsetLeft!=(document.body.scrollLeft+followObj_x)) {
     var dx=(document.body.scrollLeft+followObj_x-followObj.offsetLeft)*delta;
     dx=(dx>0?1:-1)*Math.ceil(Math.abs(dx));
     followObj.style.left=followObj.offsetLeft+dx;
     }

    if(followObj.offsetTop!=(document.body.scrollTop+followObj_y)) {
     var dy=(document.body.scrollTop+followObj_y-followObj.offsetTop)*delta;
     dy=(dy>0?1:-1)*Math.ceil(Math.abs(dy));
     followObj.style.top=followObj.offsetTop+dy;
     }
    followObj.style.display = '';
   }
  } 
  function closeBanner()
  {
   closeB=true;
   return;
  }

 var theFloaters  = new floaters();
 //
 theFloaters.addItem('followDiv1','document.body.clientWidth-90',300,'<a onClick="closeBanner();" href=http://www.cx56w.com/sjdx.asp target=_blank><img src=images/sjdx.gif border=0 ></a><br><img src=images/close.gif onClick="closeBanner();">');
 theFloaters.addItem('followDiv2',10,300,'<a onClick="closeBanner();" href=http://www.cx56w.com/sjdx.asp target=_blank><img src=images/sjdx.gif border=0 ></a><br><img src=images/close.gif onClick="closeBanner();">');
 theFloaters.play();
</script>
<!--#include file="top_fz.html"-->
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
<table width="780"  border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td height="3"></td>
</tr>
</table>

<table width="780" border="0" cellspacing="0" cellpadding="0" align="center" background="images/search_bg.gif">
  <tr> 
    <td width="88"> </td>
    <td height="30" valign="middle"> 
      <table id=Table1 cellspacing=0 cellpadding=0 width="96%" align=center 
      border=0>
        <form method="POST" action="fz.asp?fz=<%=fz%>" name="form1">
                <tr> 
                  
            <td height="30" align="left">�����أ�
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
              �߼����� | ʹ�ð���</td>
                </tr>
                <SCRIPT language=javascript>
            setup();
          </SCRIPT>
              </form>
            </table>      
	  </td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="4"></td>
  </tr>
</table>
<tr> 
  <td align="center"> 
    <TABLE cellSpacing=0 width=778 align=center border=0>
      <TR> 
        <TD vAlign=top width=280> 
          <table width="275" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td height="180" valign="top"><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
                    <TD height=24 align=middle><img src="images/dl.jpg" width="275" height="24"></TD>
</TR>
</TBODY>
</TABLE> 
                <TABLE cellSpacing=1 cellPadding=0 width="100%" bgColor=#FF9E0B 
        border=0>
                  <TR> 
                    <TD vAlign=top align=middle bgColor=#ffffff> 
                      <%if session("user")="" then%>
                      <TABLE cellSpacing=0 cellPadding=4 width="100%" border=0>
                        <TR> 
                          <TD vAlign=top background=images/login2.jpg bgColor=#f6f6f6 height=144 align="right"> 
                            <TABLE cellSpacing=0 cellPadding=0 width="68%"  border=0>
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
                                <TR align="center"> 
                                  <TD colSpan=2 height=29><a href="#">�����˱�ź����룿</a> 
                                  </TD>
                                </TR>
                                <INPUT type=hidden value=login name=send>
				  <input type="hidden" name="action" value="fz">
                                <TR> 
                                  <TD colSpan=2 height=29> 
                                    <DIV align=center> <span style="background-color:Transparent;"> 
                                      <input id="CheckBox1" type="checkbox" name="CheckBox1" checked />
                                      <label for="CheckBox1">�Ժ��Զ���¼</label></span> 
                                      <INPUT class=noo type=image 
                        src="images/soucool_006.gif" width="47" height="19">
                                      &nbsp; <span style="background-color:Transparent;"> 
                                      <label for="CheckBox1"></label></span></DIV>
                                  </TD>
                                </TR>
                                <TR align="center"> 
                                  <TD colSpan=2 height=26> <a href="Reg/User_Reg.asp"><img src="images/reg.gif" border="0"></a> 
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
                          <TD vAlign=middle background=images/login_over.jpg bgColor=#f6f6f6 height=160> 
                            <TABLE cellSpacing=0 cellPadding=0 width="40%" align=right border=0>
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
                <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" cellspacing=0 cellpadding=1 width=100% border=0>
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
          <table width="275" border="0" cellpadding="0" cellspacing="0" bgcolor="#f6f6f6" class="9p">
            <tr> 
              <td height="22" valign="top"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="275" height="85">
                  <param name="movie" value="flash/gdzs.swf" />
                  <param name="quality" value="high" />
                  <embed src="flash/gdzs.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="275" height="85">
                  </embed> 
                </object></td>
            </tr>
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
				     rs_info.open "select top 16 * from qyml where province='"&fz&"' or city='"&fz&"' order by id desc",conn,1,1
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
          <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor="#CCCCCC">
            <TR>
              <TD bgcolor="#F4FEE7" height="23" style="padding-bottom:3px" valign="bottom">&nbsp;<font color="#FF0000">��վ�ںţ�</font><marquee direction='left' scrollamount='2' width='420' onMouseOut='this.start()' onMouseOver='this.stop()'>���Ƶ���������Ϣ����������ȫ��������Ϣ���ˣ�</marquee> 
              </TD>
</TR>
</TABLE>

		  <table width="92%" border="0" height="5" cellpadding="0" cellspacing="0">
            <tr> 
              <td></td>
            </tr>
          </table>
          <%if typeselect="" then%>
          <!--#include file="middle2.asp"-->
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
          <!--#include file="file_gq2.asp"-->
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
                <!--#include file="other2.asp"-->
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
