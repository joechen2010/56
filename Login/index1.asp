<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../Member/check.asp"-->
<%
   set rs_m=conn.execute("select count(*) from message where New=0 and T_User='"&session("user")&"'")
   set rs_p=conn.execute("select count(*) from info2 where gsid='"&session("user")&"'")
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
<LINK href="../Member/images/main.css" type=text/css rel=stylesheet>
<STYLE type=text/css>.style1 {
	COLOR: #000000
}
.style2 {
	COLOR: #ffffff
}
</STYLE>
<META content="MSHTML 6.00.2462.0" name=GENERATOR></HEAD>
<BODY>
<!--#include file="../Member/head.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width=750 border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=160 bgColor=#2c68b1 height=566>
<!--#include file="../Member/left.htm"-->
      </TD>
    <TD id=main vAlign=top align=middle bgColor=#ffffff>
      <TABLE cellSpacing=0 cellPadding=3 width="94%" border=0>
        <TBODY>
        <TR>
          <TD id=pos align=left width="83%" height=38><IMG height=11 
            src="../Member/images/icon03.gif" width=15> <A 
            href="http://www.bearingpage.com/">诚信物流网</A> &gt; 会员控制中心首页</TD>
          <TD width="17%"><IMG height=32 
            src="../Member/images/icon_onlineservice.gif" 
      width=132></TD>
        </TR></TBODY></TABLE>
      <table width="94%"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="../Member/images/icon02.gif" align="absmiddle" width="27" height="19"> 
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
          <td height="31">&nbsp; 您的位置：<a href='../Member/index.asp' class='12link'>会员自助管理</a>--&gt;首页</td>
        </tr>
        <tr> 
          <td></td>
        </tr>
        
        <tr> 
          <td height="60" align="center"> <br>
            <table cellspacing=0 cellpadding=0 width="97%" border=0>
              <tbody> 
              <tr> 
                <td width=10><img height=26 src="../Member/images/title_blue_1.gif" 
            width=10></td>
                <td class=titletx align=left width=537 
          background=../Member/images/title_blue_2.gif>系统状态</td>
                <td width=8><img height=26 src="../Member/images/title_blue_3.gif" 
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
                            <td width="20%" align=center bgcolor="#F3F3F3"><a href="../Member/inbox.asp" class="12linkBlue">查看最新留言</a><a href="../Member/product_add.asp" class="12linkBlue"></a>                            </td>
                            <td>您有 <%=rs_m(0)%> 最新留言！ </td>
                          </tr>
                          </tbody> 
                        </table>                      </td>
                    </tr>
                    </tbody> 
                  </table>                </td>
              </tr>
              </tbody> 
            </table>          </td>
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
                            <td width="20%" align=center bgcolor="#F3F3F3"><a href="../Member/info_manage.asp" class="12linkBlue">我的商机列表</a><br>
                              <a href="../Member/info_add.asp" class="12linkBlue">发布供求商机</a></td>
                            <td>您已经发布<span class="red"><font color="#FF0000"><%=rs_info(0)%></font></span>条商机，其中生效的<span class="red"><font color="#FF0000"><%=rs_info1(0)%></font></span>条，未生效的<span class="red"><font color="#FF0000"><%=rs_info2(0)%></font></span>条！</td>
                          </tr>
                          </tbody> 
                        </table>                      </td>
                    </tr>
                    </tbody> 
                  </table>                </td>
              </tr>
              </tbody> 
            </table>          </td>
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
                            <td width="20%" align=center bgcolor="#F3F3F3"><a href="../Member/Pro_Manage.asp" class="12linkBlue">我的产品列表</a><br>
                              <A 
            href="<%if session("hydj")=0 then response.write "javascript:msg()" else response.write "Pro_add.asp"%>">发布公司产品</a></td>
                            <td>您已经发布<span class="red"><font color="#FF0000"><%=rs_p(0)%></font></span>个产品，其中生效的<span class="red"><font color="#FF0000"><%=rs_p1(0)%></font></span>个，未生效的<span class="red"><font color="#FF0000"><%=rs_p2(0)%></font></span>个！</td>
                          </tr>
                          </tbody> 
                        </table>                      </td>
                    </tr>
                    </tbody> 
                  </table>                </td>
              </tr>
              </tbody> 
            </table>          </td>
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
            href="<%if session("hydj")=0 then response.write "javascript:msg()" else response.write "qyry.asp"%>">发布荣誉图片</a>                            </td>
                            <td>您已经发布了<span class="red"><font color="#FF0000"><%=rs_h(0)%></font></span>张企业荣誉图片！</td>
                          </tr>
                          </tbody> 
                        </table>                      </td>
                    </tr>
                    </tbody> 
                  </table>                </td>
              </tr>
              </tbody> 
            </table>          </td>
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
                            <td width="20%" align=center bgcolor="#f3f3f3"><a href="../Member/com_edit.asp" class="12linkBlue">修改公司信息</a><br>
                              <a href="../Member/modify_info.asp" class="12linkBlue">修改联系信息</a><br>                            </td>
                            <td valign="top">会员级别： <%=ss%><a href="../About/update_esc.asp" target="_blank"><img src="../Member/images/vipdipic.gif" alt="升级易商城" width="207" height="45" border="0"></a>                            </td>
                          </tr>
                          </tbody> 
                        </table>                      </td>
                    </tr>
                    </tbody> 
                  </table>                </td>
              </tr>
              </tbody> 
            </table>          </td>
        </tr>
      </table>
      </TD></TR></TABLE>
<!--#include file="../Member/bottom.htm"--></BODY></HTML>
