<HTML>
<HEAD>
<TITLE>168������Ϣ��</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="images/default.css" type=text/css rel=stylesheet>
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
            href="index.asp">��������</a></td>
        </tr>
        <tr> 
          <td background=images/menus_s.gif height=43><a 
            href="service.asp">�������</a></td>
        </tr>
        <tr> 
          <td background=images/menus_g.gif height=43><a 
            href="guide.asp">��վָ��</a></td>
        </tr>
        <tr> 
          <td background=images/menus_m.gif height=43><a 
            href="hyzh.asp">�ʷѱ�׼</a></td>
        </tr>
        <tr> 
          <td background=images/menus_n.gif height=43><a 
            href="ad.asp">������</a></td>
        </tr>
        <tr> 
          <td background=images/menus_c.gif height=43><a 
            href="contactus.asp">��ϵ����</a></td>
        </tr>
        <tr> 
          <td background=images/menus_ag.gif height=43><a 
            href="../agent/index.asp">������ר��</a></td>
        </tr>
        </tbody> 
      </table>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
    </TD>
    <TD class=td_bd_b vAlign=top align=middle height=800>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>&nbsp;</TD>
        </TR></TBODY></TABLE>
      		  <%set rs=conn.execute("select top 1 contact from service")
		  response.Write rs("contact")
		  %>
      <TABLE cellSpacing=0 cellPadding=0 width="95%" border=0>
        <TBODY>
        <TR>
          <TD align=right><A class=t10pt 
            href="http://china.cx56w.com/aboutus/cooperate.asp#">Top</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<!--#include file="copyright.htm"--></BODY></HTML>