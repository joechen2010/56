<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<%
Dim gsid
gsid=SafeRequest("gsid",1)
if gsid="" then
    call WriteErrMsg()
end if
set rs_q=server.createobject("adodb.recordset")
sql="select * from qyml where id="&gsid&""
rs_q.open sql,conn,1,1
if rs_q.eof and rs_q.bof then
    call WriteErrMsg()
else
    dim sfilename,strUnit
    strUnit="条信息"
    sfilename="Trade.asp?gsid=" & gsid & ""
   	dim idlist
   	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if
  	dim sql
   	dim rs
%>
<HTML>
<HEAD>
<TITLE><%=rs_q("qymc")%> - 普通会员 - 供应信息</TITLE>
<STYLE type=text/css>BODY {
	FONT-SIZE: 5pt
}
TD {
	FONT-SIZE: 12px; COLOR: #174b03; FONT-FAMILY: 宋体
}
.f14 {
	FONT-SIZE: 14px; LINE-HEIGHT: 120%
}
.noline:link {
	COLOR: #000; TEXT-DECORATION: none
}
.noline:visited {
	COLOR: #000; TEXT-DECORATION: none
}
.noline:hover {
	COLOR: #f60; TEXT-DECORATION: none
}
</STYLE>
<SCRIPT language=JavaScript type=text/JavaScript>
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</SCRIPT>
<META content="MSHTML 6.00.2462.0" name=GENERATOR></HEAD>
<BODY oncontextmenu="return false" onselectstart="return false" 
ondragstart="return false" vLink=#174b03 aLink=#174b03 link=#174b03 topMargin=0 
onload="MM_preloadImages('images/MB_18.gif','images/MB_20.gif','images/MB_22.gif','images/MB_24.gif')">
<!--#include file="head.htm"-->
<TABLE height=373 cellSpacing=0 cellPadding=0 width=778 align=center 
bgColor=#ffffff border=0>
  <TBODY> 
  <TR> 
    <TD colSpan=2 height=30> 
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY> 
        <TR> 
          <TD width=168><IMG height=30 src="images/MB_12.gif" width=168></TD>
          <TH align=middle width=586 background=images/MB_17.gif><FONT 
            color=#ffffff size=3><%=rs_q("qymc")%></FONT></TH>
          <TD align=right width=24 background=images/MB_17.gif><IMG height=30 src="images/MB_19.gif" width=12></TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD>
  </TR>
  <TR> 
    <TD vAlign=top width=167 background=images/MB_38.gif><IMG height=123 src="images/MB_21.gif" width=167>
	  <A onmouseover="MM_swapImage('Image16','','images/MB_18.gif',1)" onmouseout=MM_swapImgRestore()       href="aboutus.asp?gsid=<%=gsid%>"><IMG height=34 src="images/MBh_18.gif" width=167 border=0 name=Image16></A>
	  <A href="trade.asp?gsid=<%=gsid%>"><IMG height=34 src="images/MB_20.gif" width=167 border=0></A>
	  <A onmouseover="MM_swapImage('Image18','','images/MB_22.gif',1)" onmouseout=MM_swapImgRestore() 
      href="message.asp?gsid=<%=gsid%>"><IMG height=34 src="images/MBh_22.gif" width=167 border=0 name=Image18></A>
	  <A onmouseover="MM_swapImage('Image19','','images/MB_24.gif',1)" onmouseout=MM_swapImgRestore() 
      href="contactus.asp?gsid=<%=gsid%>"><IMG height=34 src="images/MBh_24.gif" width=167 border=0 name=Image19></A> </TD>
    <TD vAlign=top width=611 rowSpan=2> 
      <TABLE cellSpacing=0 cellPadding=0 width="83%" align=right border=0>
        <TBODY> 
        <TR> 
          <TD>&nbsp;</TD>
        </TR>
        <TR> 
          <TD><IMG height=12 src="images/MB_26.gif" width=601></TD>
        </TR>
        <TR> 
          <TD background=images/MB_28.gif height=35> 
            <TABLE cellSpacing=0 cellPadding=0 width="95%" align=center border=0>
              <TBODY> 
              <TR> 
                <TD><IMG height=26 src="images/syjh.gif" width=341></TD>
              </TR>
              <TR> 
                <TD height=10></TD>
              </TR>
              </TBODY> 
            </TABLE>
          </TD>
        </TR>
        <TR> 
          <TD background=images/MB_28.gif height=12> <%
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
            <table width="90%" border="0" cellspacing="1" cellpadding="4" align="center">
              <tr bgcolor="#43ace0"> 
                <td width="24%" height="28" align="center"><span 
                        class=login>图片</span></td>
                <td width="37%" align="center"><span 
                    class=login>标题</span></td>            
                <td width="21%" align="center"><span 
                    class=login>发布日期</span></td>
                <td width="18%" align="center"><span 
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
              <tr> 
                <td width="24%" align="center" bgcolor="<%=bgcolor%>"> 
                  <%
		if rs("ProImg")="" or isnull(rs("ProImg")) then
		   response.write "<img src=""../trade/images/nopic.gif"" width=""64"" height=""64"" border=1 bordercolor=#cccccc>"
		else
		   response.write "<img src=""../InfoUpLoad/"&rs("ProImg")&""" width=""64"" height=""64"" border=1 bordercolor=#cccccc>"
		end if
		%>
                </td>
                <td width="37%" bgcolor="<%=bgcolor%>">【<%=rs("infotype")%>】<a href="../trade/InfoView.asp?InfoID=<%=rs("InfoID")%>" target="_blank"><%=rs("title")%></a></td>
                <td width="21%" align="center" bgcolor="<%=bgcolor%>"><%=rs("AddTime")%></td>
                <td width="18%" align="center" bgcolor="<%=bgcolor%>"><%=rs("AddTime")+rs("validate")%></td>
    
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
          </TD>
        </TR>
        <TR> 
          <TD background=images/MB_28.gif height=13><TABLE cellSpacing=0 cellPadding=0 width="95%" align=center border=0>
              <TBODY> 
              <TR> 
                <TD><IMG height=26 src="images/lxwm.gif" width=345></TD>
              </TR>
              <TR> 
                <TD height=10></TD>
              </TR>
              </TBODY> 
            </TABLE></TD>
        </TR>
        <TR> 
          <TD background=images/MB_28.gif height=25> <BR>
            <table width="90%" border="0" cellpadding="5" cellspacing="0" class=S2 style="border-top:1px dotted #FF6600" align="center">
              <tr bgcolor="#f6f6f6"> 
                <td width="51" valign="top" nowrap class=S2>地　址：</td>
                <td width="482" colspan="3" class=S2><%=rs_q("address")%></td>
              </tr>
              <tr> 
                <td class=S2 width="51">邮　编：</td>
                <td class=S2 colspan="3" width="482"><font face="Arial"><%=rs_q("post")%></font></td>
              </tr>
              <tr bgcolor="#f6f6f6"> 
                <td class=S2 width="51">电　话：</td>
                <td class=S2 colspan="3" width="482"> 
                  <%response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3") %>
                </td>
              </tr>
              <tr> 
                <td class=S2 width="51">传　真：</td>
                <td class=S2 colspan="3" width="482"> 
                  <% response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
                </td>
              </tr>
              <tr bgcolor="#f6f6f6"> 
                <td class=S2 width="51"><font face="Arial">E-mail：</font></td>
                <td class=S2 colspan="3" width="482"><font face="Arial"><a href="mailto:<%=rs_q("email")%>"><%=rs_q("email")%></a></font></td>
              </tr>
              <tr> 
                <td class=S2 width="51">主　页：</td>
                <td class=S2 colspan="3" width="482"><font face="Arial"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a> 
                  </font></td>
              </tr>
            </table>
            <BR>
          </TD>
        </TR>
        <TR> 
          <TD><IMG height=41 src="images/MB_31.gif" width=601></TD>
        </TR>
        <TR> 
          <TD>&nbsp;</TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD>
  </TR>
  <TR> 
    <TD vAlign=top align=middle width=167 background=images/MB_38.gif height="100%"><BR>
      <BR>
    </TD>
  </TR>
  </TBODY> 
</TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <TBODY> 
  <TR> 
    <TD bgColor=#98a2c9 height=2></TD>
  </TR>
  <TR> 
    <TD> 
      <!--#include file="Copyright.htm"-->
    </TD>
  </TR>
  </TBODY> 
</TABLE>
</BODY></HTML>
<%
	rs_q.close
	set rs_q=nothing
	conn.close
	set conn=nothing
	end if
	%>