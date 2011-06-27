<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/function.asp"-->
<%
Dim GsID
GsID=SafeRequest("GsID",1)
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
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 首页</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
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

<table width="160" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="20">&nbsp;</td>
          <td width="120">&nbsp;</td>
          <td width="20">&nbsp;</td>
        </tr>
        <tr> 
          <td background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td background="images/bk_top2.jpg"> <DIV class=tile2 
              align=center><STRONG> 联系我们</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr> 
          <td  width="160" colspan="3" valign="top" nowrap class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999; border-top: 1px solid #999999;">&nbsp;</td>
        </tr>
        <tr bgcolor="#F6F6F6"> 
          <td width="160" height="25" colspan="3" nowrap class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">地址：<%=rs_q("address")%></td>
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
          <td width="20" background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td width="731" background="images/bk_top2.jpg"><DIV class=tile2 
                      ><STRONG>公司简介</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr align="center"> 
          <td colspan="3" valign="top" style="border: 1px solid #999999;"> 
            <table width="98%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td height="25"  style="FONT-SIZE: 10pt; LINE-HEIGHT: 23px"> 
                  <%if rs_q("qyjj")<>"" then 
        response.write Replace(rs_q("qyjj"),chr(13),"<br>&nbsp;&nbsp;&nbsp;")
	else
	response.write rs_q("qyjj")
end if
%>
                </td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td height="30" colspan="3">&nbsp;</td>
        </tr>
        <tr> 
          <td background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td background="images/bk_top2.jpg"><DIV class=tile2 
                      ><STRONG>产品展示</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr> 
          <td colspan="3" style="border: 1px solid #999999;"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
  <%set rs_p=server.createobject("adodb.recordset")
			  sql="select * from spzs where flag=1 and gsid='"&rs_q("user")&"' Order By AddTime Desc"
			  rs_p.open sql,conn,1,1
			  if rs_p.eof and rs_p.bof then
			      response.write "<td>暂无产品发布！</td>"  
			  else
			  dim i
			  i=0
			    do while not rs_p.eof
			  %>
                <td align="center" valign="top"> 
                  <TABLE cellSpacing=2 cellPadding=2 width="98%" align=center 
                  border=0>
                    <TBODY>
                    <TR>
                      <TD align=middle> 
                        <%if rs_p("ProIMG")<>"" then%>
                        <a href="../../Product/ProDetail.asp?gsid=<%=gsid%>&ProID=<%=rs_p("ProID")%>" TARGET="_blank"> 
                        <IMG  src="../../ProUpload/<%=rs_p("ProIMG")%>" width=72 border=1 height=72 style="border: 1px solid #000000"></a> 
                        <%else%>
                        <img src="../../Product/images/nopic.gif" width="72" height="72" border=1 style="border: 1px solid #000000"> 
                        <%end if%>
                      </TD>
                    </TR>
                    <TR>
                      <TD class=style>
                        <DIV align=center><a href="../../Product/ProDetail.asp?gsid=<%=gsid%>&ProID=<%=rs_p("ProID")%>" TARGET="_blank"><%=rs_p("BearinNo")%></a></DIV>
                      </TD></TR></TBODY></TABLE>
                </td>						
                <%
						i=i+1
	                     if i>=4 then exit do
	                    rs_p.movenext
	                   loop
					   end if 
					   rs_p.close
					   set rs_p=nothing
						%>
  </tr>
</table>
          </td>

        </tr>
        <tr> 
          <td height="30" colspan="3">&nbsp;</td>
        </tr>
        <tr> 
          <td background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td background="images/bk_top2.jpg"><DIV class=tile2 
                      ><STRONG>供求信息</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr> 
          <td colspan="3" align="center" style="border: 1px solid #999999;">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><TABLE class=style cellSpacing=2 cellPadding=2 width="100%" 
            align=center bgColor=#f6f6f6 border=0>
                    <TBODY>
                      <TR> 
                        <TD>最新供求：</TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <TABLE height=43 cellSpacing=2 cellPadding=0 width="100%" 
            align=center border=0>
                    <TBODY>
                      <TR bgColor=#eef0fd> 
                        <%set rs_info=server.createobject("adodb.recordset")
			  sql="select * from info where flag=1 and gsid='"&rs_q("user")&"' Order By AddTime Desc"
			  rs_info.open sql,conn,1,1
			  if rs_info.eof and rs_info.bof then
			      response.write "<td>暂无信息！</td>"
			  else
			  dim n
			  n=0
			  do while not rs_info.eof
			  %>
                        <TD width="49%"> <table class=style height=80 cellspacing=0 cellpadding=0 
                  width="95%" align=center border=0>
                            <tbody>
                              <tr> 
                                <td width="20%"> <div align=center>【<font 
color=#ff0000><%=rs_info("infotype")%></font>】</div></td>
                                <td width="63%"><a href="../../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>"><%=rs_info("Title")%></a> 
                                  <span class=zi2>(<%=rs_info("AddTime")%>)</span> 
                                  <br> <a href="../../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>">[<font color="#FF6600">详细信息</font>]</a><br> 
                                </td>
                                <td width="17%"> <div align=center><font color=#000000><img height=32 
                        src="images/messenger_icon_offline.gif" 
                        width=38><br>
                                    <a 
                        href="message.asp?gsid=<%=gsid%>">询价</a></font></div></td>
                              </tr>
                            </tbody>
                          </table></TD>
                        <%
			  n=n+1
	          if n>=4 then exit do
			  if n mod 2 =0 then
			     response.write "</tr><TR bgColor=#eef0fd>"
			  end if
	          rs_info.movenext
	          loop
			  end if
			  rs_info.close
			  set rs_info=nothing
			  %>
                      </TR>
                    </TBODY>
                  </TABLE></td>
              </tr>
            </table>
</td>
        </tr>
        <tr> 
          <td colspan="3">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="Copyright.htm" -->
<%END IF%>
</body>
</html>