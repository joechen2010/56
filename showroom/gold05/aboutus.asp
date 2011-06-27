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
          <td class=font14 width="92%"><strong>公司简介</strong></td>
          <td align=right width="4%"><img height=29 src="images/img_33.gif" 
            width=10></td>
        </tr>
        </tbody> 
      </table>
      <table height=200 cellspacing=0 cellpadding=0 width="96%" border=0>
        <tbody> 
        <tr> 
          <td> 
            <table align=right border=0 cellpadding=4 cellspacing=4>
              <tbody> 
              <tr> 
                <td> 
                  <%if rs_q("cimg")<>"" then
		     response.write "<img src=""../../Member/"&rs_q("cimg")&""" height=180>"
		  else%>
                  <img 
            src="images/nopicture.gif" width=260> 
                  <%end if%>
                </td>
              </tr>
              </tbody> 
            </table>
            <%if rs_q("qyjj")<>"" then 
        response.write Replace(rs_q("qyjj"),chr(13),"<br>&nbsp;&nbsp;&nbsp;")
	else
	response.write rs_q("qyjj")
end if
%>
          </td>
        </tr>
        </tbody> 
      </table>
      <table width="96%" border="0" cellspacing="1" cellpadding="4">
        <tr> 
          <td bgcolor="#fff4e5" width="155"> 
            <div align="right">主营产品或服务： </div>
          </td>
          <td width="585" bgcolor="#F7F7F7"><span class=other><%=rs_q("Zycp")%></span></td>
        </tr>
        <tr> 
          <td bgcolor="#fff4e5" width="155"> 
            <div align="right">企业经营性质：</div>
          </td>
          <td width="585" bgcolor="#F7F7F7"><span class=other><%=rs_q("qyxz")%></span></td>
        </tr>
        <tr> 
          <td bgcolor="#fff4e5" width="155"> 
            <div align="right">经营模式： </div>
          </td>
          <td width="585" bgcolor="#F7F7F7"><span 
            class=other><%=rs_q("Qylb")%></span></td>
        </tr>
        <tr> 
          <td bgcolor="#fff4e5" height="31" width="155"> 
            <div align="right">法人代表/企业负责人：</div>
          </td>
          <td height="31" width="585" bgcolor="#F7F7F7"><span class=other><%=rs_q("Frdb")%></span></td>
        </tr>
        <tr> 
          <td bgcolor="#fff4e5" width="155"> 
            <div align="right">公司注册地： </div>
          </td>
          <td width="585" bgcolor="#F7F7F7"><%=rs_q("province")%><%=rs_q("city")%><%=rs_q("Area")%> 
          </td>
        </tr>
        <tr> 
          <td bgcolor="#fff4e5" width="155"> 
            <div align="right">公司经营地：</div>
          </td>
          <td width="585" bgcolor="#F7F7F7"><span 
            class=other><%=rs_q("address")%></span></td>
        </tr>
        <tr> 
          <td bgcolor="#fff4e5" width="155"> 
            <div align="right">注册资金：</div>
          </td>
          <td width="585" bgcolor="#F7F7F7"><span class=other><%=rs_q("Zczj")%> 
            万元</span></td>
        </tr>
        <tr> 
          <td bgcolor="#fff4e5" width="155"> 
            <div align="right">员工人数：</div>
          </td>
          <td width="585" bgcolor="#F7F7F7"><span class=other><%=rs_q("Ygrs")%></span></td>
        </tr>
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
<%end if%>
