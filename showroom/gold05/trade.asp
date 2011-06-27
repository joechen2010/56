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
<LINK href="images/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<table cellspacing=0 cellpadding=0 width="100%" border=0>
  <tbody> 
  <tr> 
    <td valign=top width=189 bgcolor=#fcf6e8> 
      <table height=69 cellspacing=0 cellpadding=0 width="100%" 
      background=images/img_bg_9.gif border=0>
        <tbody> 
        <tr> 
          <td valign="bottom"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="20"> 
                  <div align="center"><font color=#ff0000>物流通会员：创建于<%=rs_q("idate")%></font></div>
                </td>
              </tr>
            </table>
          </td>
          <td align=right width=10><img height=69 src="images/img_11.gif" 
            width=10></td>
        </tr>
        </tbody> 
      </table>
      <!--#include file="left.htm"-->
      <table cellspacing=0 cellpadding=0 width=100% border=0>
        <tbody> 
        <tr> 
          <td align=middle height="15">&nbsp;</td>
        </tr>
        <tr> 
          <td align=middle> 
            <table width="96%" border="0" cellpadding="0" cellspacing="0" class=S2 style="border-top:1px dotted #FF6600" align="center">
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
    </td>
    <td valign=top align=middle> 
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
        <tbody> 
        <tr> 
          <td valign=top height=12><img height=10 src="images/img_12.gif" 
            width=11></td>
          <td valign=top align=right><img height=10 src="images/img_14.gif" 
            width=10></td>
        </tr>
        </tbody> 
      </table>
      <table cellspacing=0 cellpadding=0 width="96%" 
      background=images/img_bg_31.gif border=0>
        <tbody> 
        <tr> 
          <td width="4%" height=29><img height=29 src="images/img_30.gif" 
            width=13></td>
          <td class=font14 width="92%"><strong>供求商机</strong></td>
          <td align=right width="4%"><img height=29 src="images/img_33.gif" 
            width=10></td>
        </tr>
        </tbody> 
      </table>
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
        <tbody> 
        <tr> 
          <td width="25%"> <br>
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
            <table width="96%" border="0" cellspacing="1" cellpadding="4" align="center">
              <tr bgcolor="#FFCC00"> 
                <td width="20%" height="28" align="center"><span 
                        class=login>图片</span></td>
                <td width="40%" align="center"><span 
                    class=login>标题</span></td>
                <td width="20%" align="center"><span 
                    class=login>发布日期</span></td>
                <td width="20%" align="center"><span 
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
              <tr bgcolor="#F7F7F7"> 
                <td width="20%" align="center"> 
                  <%
		if rs("ProImg")="" or isnull(rs("ProImg")) then
		   response.write "<img src=""../../trade/images/nopic.gif"" width=""72"" height=""72"" border=1 bordercolor=#cccccc>"
		else
		   response.write "<img src=""../../InfoUpLoad/"&rs("ProImg")&""" width=""72"" height=""72"" border=1 bordercolor=#cccccc>"
		end if
		%>
                </td>
                <td width="40%">【<%=rs("infotype")%>】<a href="../../trade/InfoView.asp?InfoID=<%=rs("InfoID")%>" target="_blank"><%=rs("title")%></a></td>
                <td width="20%" align="center"><%=rs("AddTime")%></td>
                <td width="20%" align="center"><%=rs("AddTime")+rs("validate")%></td>
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
        </tbody> 
      </table>
    </td>
    <td valign=top width=15 bgcolor=#fcf6e8> 
      <table height=69 cellspacing=0 cellpadding=0 width="100%" 
      background=images/img_bg_9.gif border=0>
        <tbody> 
        <tr> 
          <td>&nbsp;</td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<!--#include file="Copyright.htm"-->
</BODY></HTML>
				  <%
	rs_q.close
	set rs_q=nothing
	conn.close
	set conn=nothing
	end if
	%>