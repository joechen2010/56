<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/function.asp"-->
<%
Dim GsID
GsID=SafeRequest("GsID",1)
if GsID="" then
     call WriteErrMsg2()
end if
set rs_q=server.createobject("adodb.recordset")
sql="select * from qyml where id="&GsID&""
rs_q.open sql,conn,1,1
if rs_q.eof and rs_q.bof then
     call WriteErrMsg2()
else
%>
<HTML>
<HEAD>
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 资信档案</TITLE>
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
          <TD class=zi2 width=10>&nbsp;</TD></TR></TBODY></TABLE><BR>
      <TABLE height=37 cellSpacing=0 cellPadding=0 width="90%" align=center 
      background=images/title.jpg border=0>
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
              align=center><STRONG>资信档案</STRONG></DIV></TD></TR></TBODY></TABLE></TD>
          <TD width="72%">&nbsp;</TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>
            <table class=style height=30 cellspacing=0 cellpadding=0 width="90%" 
            align=center bgcolor=#ffffff background=images/bj.jpg 
            border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ffddbb><span class=zi1>　会员活动记录 会员服务第 1 年</span></td>
              </tr>
              </tbody>
            </table>
            <table class=style cellspacing=1 cellpadding=6 width="90%" 
            align=center border=0>
              <tr bgcolor=#f5f5f5> 
                <td width="24%" bgcolor=#faffdd align="left">已在诚信物流网购买的服务</td>
                <td width="61%" bgcolor=#fffbec> <font color="#FF0000"><%
				   if rs_q("flag")=0 then
	  ss="免费会员"
   elseif rs_q("flag")=1 then
	  ss="铜牌会员"
   elseif rs_q("flag")=2 then
	  ss="银牌会员"
   elseif rs_q("flag")=3 then
	  ss="金牌会员"
   end if
   response.write ss
				%></font>
                  (<%=rs_q("BeginDate")%> ~ <%=rs_q("enddate")%>)</td>
              </tr>
              <tbody> 
              <tr bgcolor=#f5f5f5> 
                <td width="24%" bgcolor=#faffdd align="left"> 注册会员的时间</td>
                <td width="61%" bgcolor=#fffbec><%=rs_q("idate")%></td>
              </tr>
              <tr bgcolor=#f5f5f5> 
                <td bgcolor=#faffdd> 最近登录的时间</td>
                <td bgcolor=#fffbec><%=rs_q("LastLoginTime")%></td>
              </tr>
              <tr bgcolor=#f5f5f5>
                <td bgcolor=#faffdd align="left" rowspan="2"> 已发布上网的信息数量</td>
                <td bgcolor=#fffbec>供求商机：<font color=#ff0000>&nbsp; </font><font color=#ff0000><%=conn.execute("select count(*) from info where gsid='"&rs_q("user")&"'")(0)%></font></td>
              </tr>
              <tr bgcolor=#f5f5f5> 
                <td bgcolor=#fffbec>展示产品：<font color=#ff0000> 
                  <%set rs_p=conn.execute("select count(*) from spzs where gsid='"&rs_q("user")&"'")
				  response.write rs_p(0)
				  rs_p.close
				  set rs_p=nothing%>
                  </font></td>
              </tr>
              </tbody> 
            </table>
            <table class=style height=30 cellspacing=0 cellpadding=0 width="90%" 
            align=center bgcolor=#fafafa border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ffddbb><span 
            class=zi1>　证书及荣誉</span></td>
              </tr>
              </tbody>
            </table>
                <table class=style cellspacing=2 cellpadding=2 width="90%" 
            align=center border=0>
                  <tbody>
                    <tr bgcolor=#f5f5f5> 
                      <td width="49%" height="89" align=middle bgcolor="#fefaed"> 
                        <%set QyzsRs=Conn.Execute("Select * from qyzs where gsid='"&rs_q("user")&"'")
				   if QyzsRs.eof and QyzsRs.bof then
				       response.write "该公司还未发布企业证书!"
				   else
				   %>
                        <table class=style cellspacing=3 cellpadding=3 width="100%" 
                  border=0>
                          <tbody>
                            <tr> 
                              <%
					dim i
					i=0
					do while not QyzsRs.eof%>
                              <td width='50%'><img src='../../Member/<%=QyzsRs("zsurl")%>'></td>
                              <%
					  i=i+1
					  if i mod 2=0 then Response.Write "</tr><tr>"
					  QyzsRs.movenext
					  loop
					  %>
                            </tr>
                          </tbody>
                        </table>
                        <%
				   end if
				   QyzsRs.close
				   set QyzsRs=nothing
				  %>
                      </td>
                    </tr>
                  </tbody>
                </table>
                <table class=style cellspacing=0 cellpadding=4 width="90%" 
            align=center border=0>
              <tbody> 
              <tr> 
                <td 
                  bgcolor=#faffdd>请注意：<br>
                  诚信物流网仅列示上述信息，上述报告仅代表报告出具日的情况，并不保证您使用上述信息时的情况与之相符，也不担保该信息的准确性，完整性和及时性，也不承担浏览者的任何商业风险。因此，诚信物流网不承担您因此而发生或致使的任何损害。诚信物流网保留全部或部分删除上述报告的权利。上述页面内容禁止以各种形式被全部或部分复制使用。 
                </td>
              </tr>
              </tbody>
            </table>
            <br></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<!--#include file="Copyright.htm"-->
                  <map name="Map"> 
                    <area shape="rect" coords="110,63,221,89" href="../../bearingpass" target="_blank">
                    <area shape="rect" coords="110,99,221,125" href="../../About/update_esc.asp" target="_blank">
                  </map>
				  </BODY></HTML>
<%end if%>
