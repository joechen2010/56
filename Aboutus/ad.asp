<HTML>
<HEAD>
<TITLE>168货运信息网</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/default.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2462.0" name=GENERATOR>
<!--#include file="../inc/conn.asp"-->
</HEAD>
<BODY>
<!--#include file="../inc/top1.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width=100% bgColor=#ffffff border=0>
  <TBODY> 
  <TR>
    <TD vAlign=top width=173 bgColor=#e7e7e7>
      <TABLE id=main_menus cellSpacing=0 cellPadding=0 width=173 border=0>
        <TBODY> 
        <TR> 
          <TD background=images/menus_a.gif height=43><A 
            href="index.asp">关于我们</A></TD>
        </TR>
        <TR> 
          <TD background=images/menus_s.gif height=43><A 
            href="service.asp">服务务款</A></TD>
        </TR>
        <TR> 
          <TD background=images/menus_g.gif height=43><A 
            href="guide.asp">网站指南</A></TD>
        </TR>
        <TR> 
          <TD background=images/menus_m.gif height=43><A 
            href="hyzh.asp">资费标准</A></TD>
        </TR>
        <TR> 
          <TD background=images/menus_n.gif height=43><A 
            href="ad.asp">网络广告</A></TD>
        </TR>
        <TR> 
          <TD background=images/menus_c.gif height=43><a 
            href="contactus.asp">联系我们</a></TD>
        </TR>
        <TR> 
          <TD background=images/menus_ag.gif height=43><A 
            href="../agent/index.asp">代理商专区</A></TD>
        </TR>
        </TBODY>
      </TABLE>
      <P>&nbsp;</P>
      <P>&nbsp;</P></TD>
    <TD class=td_bd_b vAlign=top align=middle height=800>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>&nbsp;</TD>
        </TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="90%" border=0>
        <TBODY>
        <TR>
          <TD><IMG height=29 src="images/ad.jpg" width=195></TD>
        </TR>
        <TR>
          <TD bgColor=#000000 height=3></TD></TR>
        <TR>
          <TD height=20>&nbsp;</TD></TR></TBODY></TABLE>
      
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD height=20>
		  		  <%set rs=conn.execute("select top 1 ads from service")
		  response.Write rs("ads")
		  %>
		  </TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<!--#include file="copyright.htm"--></BODY></HTML>
