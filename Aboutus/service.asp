<HTML>
<HEAD>
<TITLE>168货运信息网</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/default.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2462.0" name=GENERATOR></HEAD>
<!--#include file="../inc/conn.asp"-->
<BODY>
<!--#include file="../inc/top1.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width=100% bgColor=#ffffff border=0>
  <TBODY> 
  <TR>
    <TD vAlign=top width=173 bgColor=#e7e7e7> 
      <table id=main_menus cellspacing=0 cellpadding=0 width=173 border=0>
        <tbody> 
        <tr> 
          <td background=images/menus_a.gif height=43><a 
            href="index.asp">关于我们</a></td>
        </tr>
        <tr> 
          <td background=images/menus_s.gif height=43><a 
            href="service.asp">服务务款</a></td>
        </tr>
        <tr> 
          <td background=images/menus_g.gif height=43><a 
            href="guide.asp">网站指南</a></td>
        </tr>
        <tr> 
          <td background=images/menus_m.gif height=43><a 
            href="hyzh.asp">资费标准</a></td>
        </tr>
        <tr> 
          <td background=images/menus_n.gif height=43><a 
            href="ad.asp">网络广告</a></td>
        </tr>
        <tr> 
          <td background=images/menus_c.gif height=43><a 
            href="contactus.asp">联系我们</a></td>
        </tr>
        <tr> 
          <td background=images/menus_ag.gif height=43><a 
            href="../agent/index.asp">代理商专区</a></td>
        </tr>
        </tbody> 
      </table>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
    </TD>
    <TD vAlign=top align=middle height=800>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>&nbsp;</TD>
        </TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="90%" border=0>
        <TBODY>
        <TR>
          <TD><IMG height=29 src="images/se.jpg" width=195></TD>
        </TR>
        <TR>
          <TD bgColor=#000000 height=3></TD></TR></TBODY></TABLE>
      <TABLE class=t9pt_b cellSpacing=0 cellPadding=0 width="90%" border=0>
        <TBODY> 
        <TR> 
          <TD align=left vAlign=top>		  		  <%set rs=conn.execute("select top 1 charges from service")
		  response.Write rs("charges")
		  %></TD>
        </TR>
        </TBODY>
      </TABLE>
      </TD></TR></TBODY></TABLE>
<!--#include file="copyright.htm"--></BODY></HTML>
