<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/function.asp"-->
<%
Dim gsid
gsid=SafeRequest("GsID",1)
if gsid="" then
    call WriteErrMsg2()
end if
set rs_q=server.createobject("adodb.recordset")
sql="select * from qyml where id="&gsid&""
rs_q.open sql,conn,1,1
if rs_q.eof and rs_q.bof then
    call WriteErrMsg2()
else
    dim sfilename,strUnit
    strUnit="条信息"
    sfilename="products.asp?gsid=" & gsid & ""
  	dim sql
   	dim rs
%>
<HTML>
<HEAD>
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 产品信息</TITLE>
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
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD align="center"><BR>
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
            <table width="90%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><%
dim rssum,maxpage,thepages,viewpage,i
maxpage=10	
	sql="select * from spzs where gsid='"&rs_q("User")&"' and flag=1 order by AddTime desc"
	Set Rs=server.createobject("Adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	    response.write "<p align=""center"">没有内容</p>"
	else
	    rssum=rs.recordcount
        if rssum mod maxpage > 0 then
            thepages=rssum\maxpage+1
        else
            thepages=rssum\maxpage
        end if

        viewpage=trim(request("page"))
        if viewpage="" or isnull(viewpage) or not(isnumeric(viewpage)) then
            viewpage=1
        elseif cint(viewpage)<1 then
            viewpage=1
        else
            if cint(viewpage)>thepages then
                viewpage=1
            end if
       end if
       i=1
       if viewpage<>"1" then
           rs.move (viewpage-1)*maxpage
       end if
  %>
                  <table width="90%" border="0" cellspacing="1" cellpadding="4" align="center">
                    <tr> 
                      <td width="20%" height="28" background="../../trade/images/index_meun2.jpg" align="center"><span 
                        class=login>图片</span></td>
                      <td width="20%" background="../../trade/images/index_meun2.jpg" align="center">产品型号</td>
                      <td width="20%" background="../../trade/images/index_meun2.jpg" align="center">数量</td>
                      <td width="20%" background="../../trade/images/index_meun2.jpg" align="center"><span 
                    class=login>现货地点</span></td>
                      <td width="20%" background="../../trade/images/index_meun2.jpg" align="center"><span 
                    class=login>发布日期</span></td>
                    </tr>
                    <%       for i=1 to maxpage
           if rs.eof then exit for
 if i mod 2=1 then
    bgcolor="#f7f7f7"
 else
    bgcolor="#ffffff"
 end if
 %>
                    <tr bgcolor="<%=bgcolor%>"> 
                      <td width="20%" align="center"> 
                        <%
		if rs("ProImg")="" or isnull(rs("ProImg")) then
		   response.write "<img src=""../../trade/images/nopic.gif"" width=""72"" height=""72"" border=1 bordercolor=#cccccc>"
		else
		   response.write "<img src=""../../ProUpLoad/"&rs("ProImg")&""" width=""72"" height=""72"" border=1 bordercolor=#cccccc>"
		end if
		%>
                      </td>
                      <td width="20%"> 
                        <div align="center">【<%=rs("Brand")%>】<a href="../../Product/ProDetail.asp?ProID=<%=rs("ProID")%>" target="_blank"><%=rs("BearinNo")%></a></div>
                      </td>
                      <td width="20%" align="center"><%=rs("Num")%></td>
                      <td width="20%" align="center"><%=rs("NowPlace")%></td>
                      <td width="20%" align="center"><%=rs("AddTime")%></td>
                    </tr>
                    <%
           rs.movenext
       next
%>
                  </table>
<%
call showpage (strUnit,sfilename,10,"#ff0000")
end if
rs.close
set rs=nothing
%>
</td>
              </tr>
            </table>

          </TD></TR></TBODY></TABLE>
    </TD></TR></TBODY></TABLE>
<!--#include file="Copyright.htm"-->
<map name="Map"> 
                    <area shape="rect" coords="110,63,221,89" href="../../bearingpass" target="_blank">
                    <area shape="rect" coords="110,99,221,125" href="../../About/update_esc.asp" target="_blank">
                  </map>
				  </BODY>
</HTML>

<%
	rs_q.close
	set rs_q=nothing
	conn.close
	set conn=nothing
	end if
	%>