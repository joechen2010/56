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
<TITLE><%=rs_q("Zycp")%> - <%=rs_q("qymc")%> - 物流通金牌企业 - 供应信息</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta name="description" content="<%=rs_q("Zycp")%> \\ <%=rs_q("qymc")%> \\ <%=rs_q("Zycp")%>" >
<meta name="keywords" content="<%=rs_q("Zycp")%> \\ <%=rs_q("qymc")%>  \\ <%=rs_q("Zycp")%>" >
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="F4FFFE">
  <tr>
    <td width="175">&nbsp;</td>
    <td width="1" background="images/hot_line.jpg"></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td align="center" valign="top"> 
      <!--#include file="left.htm"-->
	  </td>
    <td width="1" background="images/hot_line.jpg"></td>
    <td align="center" valign="top"> 
      <table width="96%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="images/skin3_15.jpg" width="167" height="27" align="center"><DIV class=tile2 
                      align=center><STRONG>供求商机</STRONG></DIV></td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="96%" border="0" cellspacing="1" cellpadding="0" bgcolor="239687">
        <tr>
          <td bgcolor="239687" height="3"></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">
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
		   response.write "<img src=""../../trade/images/nopic.gif"" width=""64"" height=""64"" border=1 bordercolor=#cccccc>"
		else
		   response.write "<img src=""../../InfoUpLoad/"&rs("ProImg")&""" width=""64"" height=""64"" border=1 bordercolor=#cccccc>"
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
      </table>
      <h3>&nbsp;</h3>
    </td>
  </tr>
</table>
        
<!--#include file="Copyright.htm"--></BODY></HTML>
<%end if %>
