<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
Dim gsID
gsID=SafeRequest("gsid",1)
if gsID="" then
    call WriteErrMsg()
end if
set rs_q=server.createobject("adodb.recordset")
sql="select * from qyml where id="&gsid&""
rs_q.open sql,conn,1,1
if rs_q.eof and rs_q.bof then
    call WriteErrMsg()
else
%>
<HTML>
<HEAD>
<TITLE><%=rs_q("qymc")%> - 普通会员 - 公司介绍</TITLE>
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
</STYLE><script language="JavaScript" type="text/JavaScript">
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
</script>
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
	  <A href="aboutus.asp?gsid=<%=gsid%>"><IMG height=34 src="images/MB_18.gif" width=167 border=0 name=Image15></A>
	  <A onmouseover="MM_swapImage('Image17','','images/MB_20.gif',1)" onmouseout=MM_swapImgRestore() 
      href="trade.asp?gsid=<%=gsid%>"><IMG height=34 src="images/MBh_20.gif" width=167 border=0 name=Image17></A>
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
                <TD><IMG height=26 src="images/gsjj.gif" width=334></TD>
              </TR>
              <TR> 
                <TD height=10></TD>
              </TR>
              </TBODY> 
            </TABLE>
          </TD>
        </TR>
        <TR> 
          <TD background=images/MB_28.gif height=12> 
            <table class=style cellspacing=5 cellpadding=0 width="95%" 
                  align=center border=0>
              <tbody> 
              <tr> 
                <td class=S2 bgcolor=#f9f9f9> 
                 <font style="font-size:12px;line-height:15pt">
                    <%if rs_q("qyjj")<>"" then 
        response.write Replace(rs_q("qyjj"),chr(13),"<br>&nbsp;&nbsp;&nbsp;")
	else
	response.write rs_q("qyjj")
end if
%>
                  </font>
                </td>
              </tr>
              <tr> 
                <td class=S> 
                  <table cellspacing=2 cellpadding=2 width="100%" align=center border=0 style="TABLE-LAYOUT: fixed; WORD-BREAK: break-all">
                    <tbody> 
                    <tr bgcolor="#E1F9FF"> 
                      <td align=left width="50%" bgcolor=#E2F5FE height="39"><font style="font-size:12px;line-height:15pt">主营产品或服务： 
                        <br>
                        <span class=other><%=rs_q("Zycp")%></span></font></td>
                      <td align=left width="50%" bgcolor=#E1F9FF height="39"><font style="font-size:12px;line-height:15pt">企业经营性质： 
                        <br>
                        <span class=other><%=rs_q("qyxz")%> </span></font></td>
                    </tr>
                    <tr> 
                      <td background="" colspan=2 height=1></td>
                    </tr>
                    <tr> 
                      <td align=left width="50%"><font style="font-size:12px;line-height:15pt">经营模式：<span class=other><%=rs_q("Qylb")%></span></font></td>
                      <td align=left><font style="font-size:12px;line-height:15pt">法人代表/企业负责人：<span class=other><%=rs_q("Frdb")%></span></font></td>
                    </tr>
                    <tr> 
                      <td background="" colspan=2 height=1></td>
                    </tr>
                    <tr bgcolor="#E1F9FF"> 
                      <td align=left width="50%"><font style="font-size:12px;line-height:15pt">公司注册地： 
                        <br>
                        <span class=other><%=rs_q("province")%><%=rs_q("city")%><%=rs_q("area")%></span></font></td>
                      <td align=left><font style="font-size:12px;line-height:15pt">公司经营地： 
                        <br>
                        <span class=other><%=rs_q("address")%> </span></font></td>
                    </tr>
                    <tr> 
                      <td background="" colspan=2 height=1></td>
                    </tr>
                    <tr> 
                      <td width="50%"><font style="font-size:12px;line-height:15pt">员工人数：<span class=other><%=rs_q("Ygrs")%></span></font></td>
                      <td align=left><font style="font-size:12px;line-height:15pt">注册资金：<span class=other> 
                        <%=rs_q("Zczj")%> 万元</span></font></td>
                    </tr>
                    <tr> 
                      <td background="" colspan=2 height=1></td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </TD>
        </TR>
        <TR> 
          <TD background=images/MB_28.gif height=13>
            <table cellspacing=0 cellpadding=0 width="95%" align=center border=0>
              <tbody> 
              <tr> 
                <td><img height=26 src="images/lxwm.gif" width=345></td>
              </tr>
              <tr> 
                <td height=10></td>
              </tr>
              </tbody> 
            </table>
          </TD>
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
    <TD vAlign=top align=middle width=167 background=images/MB_38.gif height="100%"><BR><BR>
    </TD>
  </TR>
  </TBODY> 
</TABLE>
<!--#include file="Copyright.htm"-->
</BODY></HTML>
<%end if%>