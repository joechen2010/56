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
</HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width="16%" background=images/bj.jpg>
      <TABLE height=254 cellSpacing=0 cellPadding=0 width=250 
      background=images/bj.jpg border=0>
        <TBODY>
        <TR>
          <TD vAlign=top>
            <TABLE height=23 cellSpacing=0 cellPadding=0 width="100%" 
              border=0><TBODY>
              <TR>
                <TD class=zi2 background=images/syt_6.jpg 
                height=23></TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD vAlign=bottom background=images/l2.jpg><IMG 
                  height=222 src="images/syt_8.jpg" width=251 
                  useMap=#Map border=0> 
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                  background=images/l2.jpg border=0>
                    <TBODY>
                    <TR>
                      <TD vAlign=top> 
                        <table cellspacing=2 cellpadding=8 width=222 
                        align=right border=0>
                          <tbody> 
                          <tr> 
                            <td align=middle height="15">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td align=middle> 
                              <table width="100%" border="0" cellpadding="0" cellspacing="0" class=S2 style="border-top:1px dotted #FF6600">
                                <tr> 
                                  <td width="22%" nowrap class=S2 height="28">地址：</td>
                                  <td width="78%" colspan="3" class=S2 height="28"><%=rs_q("address")%></td>
                                </tr>
                                <tr> 
                                  <td class=S2 width="22%" height="28">邮编：</td>
                                  <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"><%=rs_q("post")%></font></td>
                                </tr>
                                <tr> 
                                  <td class=S2 width="22%" height="28">电话：</td>
                                  <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"> 
                                    <% Response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3")%>
                                    </font></td>
                                </tr>
                                <tr> 
                                  <td class=S2 width="22%" height="28">传真：</td>
                                  <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"> 
                                    <% Response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
                                    </font></td>
                                </tr>
                                <tr> 
                                  <td class=S2 width="22%" height="28">主页：</td>
                                  <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a>&nbsp; 
                                    </font></td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                          </tbody> 
                        </table>
                        <BR><BR>
                        <TABLE cellSpacing=0 cellPadding=0 width="80%" 
                        align=center border=0>
                          <TBODY>
                          <TR>
                            <TD></TD></TR></TBODY></TABLE>
                      </TD>
                    </TR>
                    <TR>
                      <TD><IMG height=18 src="images/l1.jpg" 
                        width=251></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE><BR></TD></TR></TBODY></TABLE></TD>
    <TD vAlign=top width="84%">
      <TABLE height=80 cellSpacing=0 cellPadding=0 width="100%" 
      background=images/syt_7.jpg border=0>
        <TBODY>
        <TR>
          <TD width="86%">
            <DIV class=tile>
            <DIV align=right><STRONG><%=rs_q("qymc")%></STRONG></DIV></DIV></TD>
          <TD width="14%">
            <DIV align=right>
            <OBJECT 
            codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0 
            height=80 width=200 
            classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000><PARAM NAME="movie" VALUE="images/line1.swf"><PARAM NAME="quality" VALUE="high"><PARAM NAME="wmode" VALUE="transparent">
                                            				                <embed 
            src="images/line1.swf" quality="high" 
            pluginspage="http://www.macromedia.com/go/getflashplayer" 
            type="application/x-shockwave-flash" width="200" 
            height="80"></embed></OBJECT></DIV></TD></TR></TBODY></TABLE>
      <TABLE height=23 cellSpacing=0 cellPadding=0 width="100%" bgColor=#fff4f4 
      border=0>
        <TBODY>
        <TR>
          <TD class=zi2 width="97%" height=23>
            <DIV align=right><font color=#ff0000>物流通金牌会员：创建于<%=rs_q("idate")%></font></DIV>
          </TD>
          <TD class=zi2 width=10>&nbsp;</TD></TR></TBODY></TABLE>
      <TABLE height=96 cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD height=96><BR>
            <TABLE height=37 cellSpacing=0 cellPadding=0 width="90%" 
            align=center background=images/title.jpg border=0>
              <TBODY>
              <TR>
                <TD width="6%"><IMG height=32 src="images/pic.jpg" 
                  width=38></TD>
                <TD width="22%">
                  <TABLE height=24 cellSpacing=0 cellPadding=0 width="69%" 
                  align=center bgColor=#ffffff border=0>
                    <TBODY>
                    <TR>
                      <TD>
                        <DIV class=tile2 
                      align=center><STRONG>公司简介</STRONG></DIV></TD></TR></TBODY></TABLE></TD>
                <TD width="72%">&nbsp;</TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD><BR>
                  <TABLE class=style cellSpacing=5 cellPadding=0 width="95%" 
                  align=center border=0>
                    <TBODY>
                    <TR>
                      <TD class=S2 align=middle bgColor=#f9f9f9></TD></TR>
                    <TR>
                      <TD class=S2 bgColor=#f9f9f9>
                        <P><%if rs_q("qyjj")<>"" then 
        response.write Replace(rs_q("qyjj"),chr(13),"<br>&nbsp;&nbsp;&nbsp;")
	else
	response.write rs_q("qyjj")
end if
%></P></TD></TR>
                    <TR>
                      <TD class=S>
                        <TABLE height=100 cellSpacing=1 cellPadding=1 
                        width="100%" align=center bgColor=#ffcccc border=0>
                          <TBODY>
                          <TR>
                            <TD bgColor=#ffffff>
                              <TABLE cellSpacing=0 cellPadding=5 width="100%" 
                              border=0>
                                <TBODY>
                                <TR>
                                <TD bgColor=#feedd1 height=100>
                                <TABLE class=style cellSpacing=5 cellPadding=5 
                                width="100%" align=center bgColor=#ffffff 
                                border=0>
                                <TBODY>
                                <TR>
                                        <TD height=85>
                                          <table width="100%" border="0" cellspacing="1" cellpadding="6">
                                            <tr> 
                                              <td bgcolor="#fff4e5" width="120">主营产品或服务： 
                                              </td>
                                              <td><span class=other><%=rs_q("Zycp")%></span></td>
                                            </tr>
                                            <tr> 
                                              <td bgcolor="#fff4e5">企业经营性质：</td>
                                              <td><span class=other><%=rs_q("qyxz")%></span></td>
                                            </tr>
                                            <tr> 
                                              <td bgcolor="#fff4e5">经营模式： </td>
                                              <td><span 
            class=other><%=rs_q("Qylb")%></span></td>
                                            </tr>
                                            <tr> 
                                              <td bgcolor="#fff4e5" height="31">法人代表/企业负责人：</td>
                                              <td height="31"><span class=other><%=rs_q("Frdb")%></span></td>
                                            </tr>
                                            <tr> 
                                              <td bgcolor="#fff4e5">公司注册地： </td>
                                              <td><%=rs_q("province")%><%=rs_q("city")%><%=rs_q("Area")%> 
                                              </td>
                                            </tr>
                                            <tr> 
                                              <td bgcolor="#fff4e5">公司经营地：</td>
                                              <td><span 
            class=other><%=rs_q("address")%></span></td>
                                            </tr>
                                            <tr> 
                                              <td bgcolor="#fff4e5">注册资金：</td>
                                              <td><span class=other><%=rs_q("Zczj")%> 
                                                万元</span></td>
                                            </tr>
                                            <tr> 
                                              <td bgcolor="#fff4e5">员工人数：</td>
                                              <td><span class=other><%=rs_q("Ygrs")%></span></td>
                                            </tr>
                                          </table>
                                            </TD>
                                      </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE><BR></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<!--#include file="Copyright.htm"-->
<%end if%>
<map name="Map"> 
  <area shape="rect" coords="110,63,221,89" href="../../bearingpass" target="_blank">
  <area shape="rect" coords="110,99,221,125" href="../../About/update_esc.asp" target="_blank">
</map>
