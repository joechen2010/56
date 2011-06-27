<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/function.asp"-->
<%
Dim gsid
gsid=SafeRequest("gsid",1)
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
    sfilename="Trade.asp?gsid=" & gsid & ""
  	dim sql
   	dim rs
%>
<HTML>
<HEAD>
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 供应信息</TITLE>
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
          <td  width="160" height="25" colspan="3" nowrap class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">地址：<%=rs_q("address")%></td>
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
      </table></td>
    <td width="803" align="center" valign="top"> 
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td width="731" background="images/bk_top2.jpg"> <DIV class=tile2 
              ><STRONG> 供求商机</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr align="center"> 
          <td colspan="3"  class=C vAlign=top borderColor=#cccccc style="border: 1px solid #999999;">
          <%
dim rssum,maxpage,thepages,viewpage,i
maxpage=10	
	sql="select * from info where gsid='"&rs_q("user")&"' order by AddTime desc"
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
       end if%>
            <table width="100%" border="0" cellspacing="1" cellpadding="4" align="center">
              <tr> 
                <td width="24%" height="28" background="../../trade/images/index_meun2.jpg" align="center"><span 
                        class=login>图片</span></td>
                <td width="37%" background="../../trade/images/index_meun2.jpg" align="center"><span 
                    class=login>标题</span></td>            
                <td width="21%" background="../../trade/images/index_meun2.jpg" align="center"><span 
                    class=login>发布日期</span></td>
                <td width="18%" background="../../trade/images/index_meun2.jpg" align="center"><span 
                    class=login>有效日期</span></td>
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
                <td width="24%" align="center"> 
                  <%
		if rs("ProImg")="" or isnull(rs("ProImg")) then
		   response.write "<img src=""../../trade/images/nopic.gif"" width=""72"" height=""72"" border=1 bordercolor=#cccccc>"
		else
		   response.write "<img src=""../../InfoUpLoad/"&rs("ProImg")&""" width=""72"" height=""72"" border=1 bordercolor=#cccccc>"
		end if
		%>
                </td>
                <td width="37%">【<%=rs("infotype")%>】<a href="../../trade/InfoView.asp?InfoID=<%=rs("InfoID")%>" target="_blank"><%=rs("title")%></a></td>
                <td width="21%" align="center"><%=rs("AddTime")%></td>
                <td width="18%" align="center"><%=rs("AddTime")+rs("validate")%></td>
    
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
        <tr> 
          <td height="30" colspan="3">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="Copyright.htm"--></BODY></HTML><%
	rs_q.close
	set rs_q=nothing
	conn.close
	set conn=nothing
	end if
	%>