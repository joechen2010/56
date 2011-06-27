<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/function.asp"-->
<%
Dim gsID
gsID=SafeRequest("gsid",1)
if gsID="" then
     call WriteErrMsg2()
end if
set rs_q=server.createobject("adodb.recordset")
sql="select * from qyml where id="&gsid&""
rs_q.open sql,conn,1,1
if rs_q.eof and rs_q.bof then
     call WriteErrMsg2()
else
%>
<HTML>
<HEAD>
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 公司介绍</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
<style>
a{
	font-size: 10pt;
	color: #000000;
	text-decoration: none;
	FONT-FAMILY: "宋体"
}
a:hover {
	color: #FF0000;
	text-decoration: none;
}

td {
	font-size: 10pt;
	text-decoration: blink;
}
</style>
</HEAD>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="head.htm" -->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="169" background="images/mb_01.jpg"><DIV class=tile> 　　<font color="#FF0000" size="+3"><STRONG><%=rs_q("qymc")%></STRONG></font>
            </DIV></td>
  </tr>
</table>
<table width="1003" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="200" align="center" valign="top"> 
      <!--#include file="left.htm" -->

<table width="160" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="20" height="16">&nbsp;</td>
          <td width="120">&nbsp;</td>
          <td width="20">&nbsp;</td>
        </tr>
        <tr> 
          <td height="31" background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td background="images/bk_top2.jpg"> <DIV class=tile2 
              align=center><STRONG> 联系我们</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr> 
          <td  width="160" height="17" colspan="3" valign="top" nowrap class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999; border-top: 1px solid #999999;">&nbsp;</td>
        </tr>
        <tr bgcolor="#F6F6F6"> 
          <td width="160" height="25" colspan="3" nowrap bgcolor="#F6F6F6" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">地址：<%=rs_q("address")%></td>
        </tr>
        <tr> 
          <td width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;" td>邮编：<font face="Arial"><%=rs_q("post")%></font></td>
        </tr>
        <tr bgcolor="#F6F6F6"> 
          <td width="160" height="25" colspan="3"  class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">电话：<font face="Arial"> 
            <% Response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3")%>
            </font></td>
        </tr>
        <tr> 
          <td width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">传真：<font face="Arial"> 
            <% Response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
            </font></td>
        </tr>
        <tr bgcolor="#F6F6F6"> 
          <td width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;"><font face="Arial">E-mail：</font><font face="Arial"><a href="mailto:<%=rs_q("email")%>"><%=rs_q("email")%></a></font></td>
        </tr>
        <tr> 
          <td width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;border-bottom: 1px solid #999999">主页：<font face="Arial"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a></font></td>
        </tr>
      </table>
    </td>
    <td width="803" align="center" valign="top"> 
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td width="731" background="images/bk_top2.jpg"> <DIV class=tile2 
              ><STRONG> 公司简介</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr> 
          <td colspan="3" align="center" vAlign=top borderColor=#cccccc  class=C style="border: 1px solid #999999;">
<TABLE class=style cellSpacing=5 cellPadding=0 width="95%" 
                  align=center border=0>
                    <TBODY>
                    <TR>
                      <TD class=S2 align=middle bgColor=#f9f9f9></TD></TR>
                    <TR>
                      
                  <TD height="25" bgColor=#f9f9f9 style="FONT-SIZE: 10pt; LINE-HEIGHT: 23px"> 
                    <P>
                      <%if rs_q("qyjj")<>"" then 
        response.write Replace(rs_q("qyjj"),chr(13),"<br>&nbsp;&nbsp;&nbsp;")
	else
	response.write rs_q("qyjj")
end if
%>
                    </P></TD></TR>
                    <TR>
                      <TD class=S>
                        <TABLE height=100 cellSpacing=1 cellPadding=1 
                        width="100%" align=center bgColor=#CCFFCC border=0>
                      <TBODY>
                          <TR>
                            
                          <TD bgColor=#CCCCCC> 
                            <TABLE cellSpacing=0 cellPadding=5 width="100%" 
                              border=0>
                              <TBODY>
                                <TR>
                                  <TD bgColor=#FFFFFF height=100> 
                                    <TABLE class=style cellSpacing=5 cellPadding=5 
                                width="100%" align=center bgColor=#ffffff 
                                border=0>
                                <TBODY>
                                <TR>
                                        <TD height=85>
                                          <table width="100%" border="0" cellspacing="1" cellpadding="6">
                                              <tr> 
                                                <td bgcolor="#BDEEDE" width="120">主营产品或服务： 
                                                </td>
                                              <td><span class=other><%=rs_q("Zycp")%></span></td>
                                            </tr>
                                            <tr> 
                                                <td bgcolor="#BDEEDE">企业经营性质：</td>
                                           <td><span class=other><%=rs_q("qyxz")%></span></td>
                                            </tr>
                                            <tr> 
                                                <td bgcolor="#BDEEDE">经营模式： </td>
                                              <td><span 
            class=other><%=rs_q("Qylb")%></span></td>
                                            </tr>
                                            <tr> 
                                                <td bgcolor="#BDEEDE" height="31">法人代表/企业负责人：</td>
                                              <td height="31"><span class=other><%=rs_q("Frdb")%></span></td>
                                            </tr>
                                            <tr> 
                                                <td bgcolor="#BDEEDE">公司注册地： </td>
                                              <td><%=rs_q("province")%><%=rs_q("city")%><%=rs_q("Area")%> 
                                              </td>
                                            </tr>
                                            <tr> 
                                                <td bgcolor="#BDEEDE">公司经营地：</td>
                                              <td><span 
            class=other><%=rs_q("address")%></span></td>
                                            </tr>
                                            <tr> 
                                                <td bgcolor="#BDEEDE">注册资金：</td>
                                              <td><span class=other><%=rs_q("Zczj")%> 
                                                万元</span></td>
                                            </tr>
                                            <tr> 
                                                <td bgcolor="#BDEEDE">员工人数：</td>
                                              <td><span class=other><%=rs_q("Ygrs")%></span></td>
                                            </tr>
                                          </table>
                                           </TD>
                                      </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></td>
        </tr>
        <tr> 
          <td height="30" colspan="3">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="Copyright.htm" -->
<%end if%>
</body>
</html>