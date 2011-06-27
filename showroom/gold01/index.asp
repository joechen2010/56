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
          <TD vAlign=top height="454"> 
            <TABLE height=23 cellSpacing=0 cellPadding=0 width="100%" 
              border=0><TBODY>
              <TR>
                <TD class=zi2 background=images/syt_6.jpg 
              height=23></TD>
              </TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD vAlign=bottom background=images/l2.jpg><img 
                  height=222 src="images/syt_8.jpg" width=251 usemap=#Map border=0> 
                  <table cellspacing=0 cellpadding=0 width="100%" 
                  background=images/l2.jpg border=0>
                    <tbody> 
                    <tr> 
                      <td valign=top> 
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
                        <table cellspacing=0 cellpadding=0 width="80%" 
                        align=center border=0>
                          <tbody> 
                          <tr> 
                            <td></td>
                          </tr>
                          </tbody> 
                        </table>
                      </td>
                    </tr>
                    <tr> 
                      <td><img height=18 src="images/l1.jpg" 
                      width=251></td>
                    </tr>
                    </tbody> 
                  </table>
                </TD></TR></TBODY></TABLE><BR></TD></TR></TBODY></TABLE></TD>
    <TD vAlign=top width="84%">
      <TABLE height=80 cellSpacing=0 cellPadding=0 width="100%" 
      background=images/syt_7.jpg border=0>
        <TBODY> 
        <TR>
          <TD width="86%">
            <DIV class=tile> 
              <DIV align=right><strong><%=rs_q("qymc")%></strong></DIV>
            </DIV></TD>
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
                      align=center><STRONG>产品展示</STRONG></DIV></TD></TR></TBODY></TABLE></TD>
                <TD width="72%">&nbsp;</TD></TR></TBODY></TABLE>
            <TABLE height=43 cellSpacing=0 cellPadding=0 width="90%" 
            align=center border=0>
              <TBODY>
              <tr><%set rs_p=server.createobject("adodb.recordset")
			  sql="select * from spzs where flag=1 and gsid='"&rs_q("user")&"' Order By AddTime Desc"
			  rs_p.open sql,conn,1,1
			  if rs_p.eof and rs_p.bof then
			      response.write "<td>暂无产品发布！</td>"  
			  else
			  dim i
			  i=0
			    do while not rs_p.eof
			  %><td><TABLE cellSpacing=2 cellPadding=2 width="90%" align=center 
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
                      </TD></TR></TBODY></TABLE></td>
						<%
						i=i+1
	                     if i>=4 then exit do
	                    rs_p.movenext
	                   loop
					   end if 
					   rs_p.close
					   set rs_p=nothing
						%></tr>
			  </TBODY></TABLE></TD></TR></TBODY></TABLE>
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
                <TD width="72%">&nbsp;</TD></TR></TBODY></TABLE><BR>
            <TABLE height=100 cellSpacing=1 cellPadding=1 width="90%" 
            align=center bgColor=#ffcccc border=0>
              <TBODY>
              <TR>
                <TD bgColor=#ffffff>
                  <TABLE cellSpacing=0 cellPadding=5 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD bgColor=#feedd1 height=100>
                        <TABLE class=style cellSpacing=5 cellPadding=5 
                        width="100%" align=center bgColor=#fffdee border=0>
                          <TBODY>
                          <TR>
                            <TD height=85> 
                              <P>
                                <%if rs_q("qyjj")<>"" then 
        response.write Replace(rs_q("qyjj"),chr(13),"<br>&nbsp;&nbsp;&nbsp;")
	else
	response.write rs_q("qyjj")
end if
%>
                              </P>
                            </TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
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
                      align=center><STRONG>供求信息</STRONG></DIV></TD></TR></TBODY></TABLE></TD>
                <TD width="72%">&nbsp;</TD></TR></TBODY></TABLE>
            <TABLE class=style cellSpacing=2 cellPadding=2 width="90%" 
            align=center bgColor=#f6f6f6 border=0>
              <TBODY>
              <TR>
                <TD>最新供求：</TD>
              </TR></TBODY></TABLE>
            <TABLE height=43 cellSpacing=2 cellPadding=0 width="90%" 
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
                <TD width="49%"> 
                  <table class=style height=60 cellspacing=0 cellpadding=0 
                  width="100%" align=center border=0>
                    <tbody> 
                    <tr> 
                      <td width="13%"> 
                        <div align=center>【<font 
color=#ff0000><%=rs_info("infotype")%></font>】</div>
                      </td>
                      <td width="70%"><a href="../../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>"><%=rs_info("Title")%></a> 
                        <span class=zi2>(<%=rs_info("AddTime")%>)</span> <a href="../../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>">[<font color="#FF6600">详细信息</font>]</a><br>
                      </td>
                      <td width="17%"> 
                        <div align=center><font color=#000000><img height=32 
                        src="images/messenger_icon_offline.gif" 
                        width=38><br>
                          <a 
                        href="message.asp?Gsid=<%=gsid%>">询价</a></font></div>
                      </td>
                    </tr>
                    </tbody> 
                  </table>
                </TD>
            
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
			  %>  </TR>
              </TBODY>
            </TABLE>
            <BR></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<!--#include file="Copyright.htm"-->
                  <map name="Map"> 
                    <area shape="rect" coords="110,63,221,89" href="../../bearingpass" target="_blank">
                    <area shape="rect" coords="110,99,221,125" href="../../About/update_esc.asp" target="_blank">
                  </map>
</BODY></HTML>

<%END IF%>