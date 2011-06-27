<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<%
   set rs_m=conn.execute("select count(*) from message where New=0 and T_User='"&session("user")&"'")
   set rs_p=conn.execute("select count(*) from spzs where gsid='"&session("user")&"'")
   set rs_p1=conn.execute("select count(*) from spzs where flag=1 and gsid='"&session("user")&"'")
   set rs_p2=conn.execute("select count(*) from spzs where flag=0 and gsid='"&session("user")&"'")
		   
   set rs_info=conn.execute("select count(*) from info where gsid='"&session("user")&"'")
   set rs_info1=conn.execute("select count(*) from info where flag=1 and gsid='"&session("user")&"'")
   set rs_info2=conn.execute("select count(*) from info where flag=0 and gsid='"&session("user")&"'")
   
   set rs_h=conn.execute("select count(*) from Qyzs where gsid='"&session("user")&"'")
%>
<HTML>
<HEAD>
<TITLE>诚信物流网 会员控制中心 - 物流商人的网上家园</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>BODY {
	MARGIN: 0px; BACKGROUND-COLOR: #2c68b1
}
</STYLE>
<LINK href="images/main.css" type=text/css rel=stylesheet>
<STYLE type=text/css>.style1 {
	COLOR: #000000
}
.style2 {
	COLOR: #ffffff
}
</STYLE>
<META content="MSHTML 6.00.2462.0" name=GENERATOR></HEAD>
<BODY>
<!--#include file="head.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width=778 border=0 align="center">
  <TBODY> 
  <TR>
    <TD vAlign=top width=160 bgColor=#2c68b1 height=566>
<!--#include file="left.htm"-->
      </TD>
    <TD id=main vAlign=top align=middle bgColor=#ffffff>
      <TABLE cellSpacing=0 cellPadding=3 width="94%" border=0>
        <TBODY>
        <TR>
          <TD id=pos align=left width="83%" height=38><IMG height=11 
            src="images/icon03.gif" width=15> <A 
            href="http://www.cx56w.com/">诚信物流网</A> &gt; 会员控制中心首页</TD>
          <TD width="17%"><IMG height=32 
            src="images/icon_onlineservice.gif" 
      width=132></TD>
        </TR></TBODY></TABLE>
      <table width="94%"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> </font></strong><b><span style="font-size: 10.5pt">尊敬的 
            </span><font face=Arial color="#FF6600" style="font-size: 10.5pt"><%=session("name")%></font><span style="font-size: 10.5pt">：您好！ 
            <%if session("uflag")=0 then
		       response.write "您暂时还没有通过审核!"
		  else
		        response.write "您是我们的" & ss &"!"
		  end if%>
            </span></b></td>
        </tr>
      </table>
      <table width="96%" border="0" cellpadding="0" cellspacing="0" class="12black">
        <tr> 
          <td height="31">&nbsp; 您的位置：<a href='index.asp' class='12link'>会员自助管理</a>--&gt;首页</td>
        </tr>
        <tr> 
          <td></td>
        </tr>
        <tr>
          <td height="60" align="center"> 
            <TABLE cellSpacing=0 cellPadding=0 width="97%" border=0>
              <TBODY> 
              <TR>
          <TD width=7 height=6><IMG height=27 
            src="images/con_green_1.gif" width=7></TD>
          <TD class=titletx align=left 
            background=images/con_green_2.gif>最新采购信息</TD>
          <TD width=6 height=6><IMG height=27 
            src="images/con_green_3.gif" width=9></TD></TR>
        <TR>
          <TD background=images/con_green_4.gif height=119></TD>
                <TD id="" vAlign=top align=left bgColor=#f6fdf6>
                  <table width="100%" border="0">
                    <tr> 
                      <td width="50%"> 
                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%set rs_i=conn.execute("select * from info where Flag=1 and infotype='采购' order by AddTime desc")
			if rs_i.eof and rs_i.bof then
			
			else
			   dim g
			   g=0
			   do while not rs_i.eof 
			%>
                          <tr> 
                            <td height="25"><font color=#FF6600>【<%=rs_i("infotype")%>】</font><a href="../trade/InfoView.asp?InfoID=<%=rs_i("infoID")%>" target="_blank"><%=left(rs_i("title"),12)%></a><span class="S">(<%=month(rs_i("AddTime"))%>/<%=day(rs_i("AddTime"))%>)</span></td>
                          </tr>
                          <%
			     g=g+1
			     if g>=5 then exit do
			          rs_i.movenext
				 loop
			 end if
			 rs_i.close
			 set rs_i=nothing
			 %>
                        </table>
                      </td>
                      <td> 
                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%set rs_i=conn.execute("select * from Offer order by UpdateTime desc")
			if rs_i.eof and rs_i.bof then
			
			else
			   g=0
			   do while not rs_i.eof 
			%>
                          <tr> 
                            <td height="25"><font color=#079807>【<%=rs_i("OfferType")%>】</font><a href="../trade/offer_view.asp?OfferID=<%=rs_i("OfferID")%>" target="_blank"><%=left(rs_i("Subject"),15)%></a><span class="S">(<%=month(rs_i("UpdateTime"))%>/<%=day(rs_i("UpdateTime"))%>)</span></td>
                          </tr>
                          <%
			     g=g+1
			     if g>=5 then exit do
			          rs_i.movenext
				 loop
			 end if
			 rs_i.close
			 set rs_i=nothing
			 %>
                        </table>
                      </td>
                    </tr>
                  </table>
                </TD>
          <TD background=images/con_green_8.gif></TD></TR>
        <TR>
          <TD width=7 height=6><IMG height=10 
            src="images/con_green_5.gif" width=7></TD>
          <TD background=images/con_green_6.gif></TD>
          <TD width=6 height=6><IMG height=10 
            src="images/con_green_7.gif" 
  width=9></TD></TR></TBODY></TABLE></td>
        </tr>
        <tr> 
          <td height="60" align="center"> <br>
            <table cellspacing=0 cellpadding=0 width="97%" border=0>
              <tbody> 
              <tr> 
                <td width=10><img height=26 src="images/title_blue_1.gif" 
            width=10></td>
                <td class=titletx align=left width=537 
          background=images/title_blue_2.gif>系统状态</td>
                <td width=8><img height=26 src="images/title_blue_3.gif" 
            width=8></td>
              </tr>
              </tbody>
            </table>
            <table width="97%" border=0 align="center" cellpadding="0" cellspacing=1 bgcolor=#CCCCCC>
              <tbody> 
              <tr> 
                <td bgcolor=#ffffff> 
                  <table width="100%" align="center" cellpadding=5 cellspacing=4>
                    <tbody> 
                    <tr> 
                      <td bgcolor=#f7f7f7> 
                        <table width="100%" height="25" cellpadding="0" cellspacing="0" class="12black">
                          <tbody> 
                          <tr> 
                            <td width="20%" height="25" align=center bgcolor=#D7D7D7>我的留言</td>
                            <td width="20%" align=center bgcolor="#F3F3F3"><a href="inbox.asp" class="12linkBlue">查看最新留言</a><a href="product_add.asp" class="12linkBlue"></a> 
                            </td>
                            <td>您有 <%=rs_m(0)%> 最新留言！ </td>
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
          </td>
        </tr>
        <tr> 
          <td height="90" align="center"> 
            <table width="97%" border=0 align="center" cellpadding="0" cellspacing=1 bgcolor=#CCCCCC>
              <tbody> 
              <tr> 
                <td bgcolor=#ffffff> 
                  <table width="100%" align="center" cellpadding=5 cellspacing=4>
                    <tbody> 
                    <tr> 
                      <td bgcolor=#f7f7f7> 
                        <table width="100%" height="45" cellpadding="0" cellspacing="0" class="12black">
                          <tbody> 
                          <tr> 
                            <td width="20%" height="45" align=center bgcolor=#D7D7D7>商业机会</td>
                            <td width="20%" align=center bgcolor="#F3F3F3"><a href="info_manage.asp" class="12linkBlue">我的商机列表</a><br>
                              <a href="info_add.asp" class="12linkBlue">发布供求商机</a></td>
                            <td>您已经发布<span class="red"><font color="#FF0000"><%=rs_info(0)%></font></span>条商机，其中生效的<span class="red"><font color="#FF0000"><%=rs_info1(0)%></font></span>条，未生效的<span class="red"><font color="#FF0000"><%=rs_info2(0)%></font></span>条！</td>
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
          </td>
        </tr>
        <tr> 
          <td height="90" align="center"> 
            <table width="97%" border=0 align="center" cellpadding="0" cellspacing=1 bgcolor=#CCCCCC>
              <tbody> 
              <tr> 
                <td bgcolor=#ffffff> 
                  <table width="100%" align="center" cellpadding=5 cellspacing=4>
                    <tbody> 
                    <tr> 
                      <td bgcolor=#f7f7f7> 
                        <table width="100%" height="45" cellpadding="0" cellspacing="0" class="12black">
                          <tbody> 
                          <tr> 
                            <td width="20%" height="45" align=center bgcolor=#D7D7D7>产品展示</td>
                            <td width="20%" align=center bgcolor="#F3F3F3"><a href="Pro_Manage.asp" class="12linkBlue">我的产品列表</a><br>
                              <A 
            href="<%if session("hydj")=0 then response.write "javascript:msg()" else response.write "Pro_add.asp"%>">发布公司产品</a></td>
                            <td>您已经发布<span class="red"><font color="#FF0000"><%=rs_p(0)%></font></span>个产品，其中生效的<span class="red"><font color="#FF0000"><%=rs_p1(0)%></font></span>个，未生效的<span class="red"><font color="#FF0000"><%=rs_p2(0)%></font></span>个！</td>
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
          </td>
        </tr>
        <tr> 
          <td height="60" align="center"> 
            <table width="97%" border=0 align="center" cellpadding="0" cellspacing=1 bgcolor=#CCCCCC>
              <tbody> 
              <tr> 
                <td bgcolor=#ffffff> 
                  <table width="100%" align="center" cellpadding=5 cellspacing=4>
                    <tbody> 
                    <tr> 
                      <td bgcolor=#f7f7f7> 
                        <table width="100%" height="25" cellpadding="0" cellspacing="0" class="12black">
                          <tbody> 
                          <tr> 
                            <td width="20%" height="25" align=center bgcolor=#D7D7D7>公司荣誉</td>
                            <td width="20%" align=center bgcolor="#F3F3F3"> <A 
            href="<%if session("hydj")=0 then response.write "javascript:msg()" else response.write "qyry.asp"%>">发布荣誉图片</a> 
                            </td>
                            <td>您已经发布了<span class="red"><font color="#FF0000"><%=rs_h(0)%></font></span>张企业荣誉图片！</td>
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
          </td>
        </tr>
        <tr> 
          <td height="90" align="center"> 
            <table width="97%" border=0 align="center" cellpadding="0" cellspacing=1 bgcolor=#CCCCCC>
              <tbody> 
              <tr> 
                <td bgcolor=#ffffff> 
                  <table width="100%" align="center" cellpadding=5 cellspacing=4>
                    <tbody> 
                    <tr> 
                      <td bgcolor=#f7f7f7> 
                        <table width="100%" height="65" cellpadding="0" cellspacing="0" class="12black">
                          <tbody> 
                          <tr> 
                            <td width="20%" height="45" align=center bgcolor=#D7D7D7>会员信息</td>
                            <td width="20%" align=center bgcolor="#f3f3f3"><a href="com_edit.asp" class="12linkBlue">修改公司信息</a><br>
                              <a href="modify_info.asp" class="12linkBlue">修改联系信息</a><br>
                            </td>
                            <td valign="top">会员级别： <%=ss%><a href="../About/update_esc.asp" target="_blank"><img src="images/vipdipic.gif" alt="升级易商城" width="207" height="45" border="0"></a> 
                            </td>
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
          </td>
        </tr>
      </table>
            </TD></TR></TABLE>
<!--#include file="bottom.htm"--></BODY></HTML>
