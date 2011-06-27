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
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 联系我们</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="FFFCF4">
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
              align=center><STRONG> 联系我们</STRONG></DIV></td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="96%" border="0" cellspacing="1" cellpadding="0" bgcolor="f37600">
        <tr>
          <td bgcolor="f37600" height="3"></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"> 
            <table class=style cellspacing=5 cellpadding=3 width="90%" 
            align=center border=0>
              <tbody> 
              <tr> 
                <td class=C valign=top width="23%" bgcolor=#f3f3f3>联系人：</td>
                <td class=C width="77%" bgcolor=#f3f3f3><span class=C><%=rs_q("name")%> 
                  (<%=rs_q("ch")%>)&nbsp;<font size="-1"><b> 
                  <%if rs_q("zw")="空" then
else 
   response.write rs_q("zw")
end if%>
                  </b></font></span></td>
              </tr>
              <tr> 
                <td class=C width="23%">手机：</td>
                <td class=C width="77%"><%=rs_q("mobile")%></td>
              </tr>
              <tr> 
                <td class=C width="23%" bgcolor=#f3f3f3>电话：</td>
                <td class=C width="77%" bgcolor=#f3f3f3><font face="Arial"> 
                  <% Response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3")%>
                  </font></td>
              </tr>
              <tr> 
                <td class=C width="23%">传真：</td>
                <td class=C width="77%"><font face="Arial"> 
                  <% Response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
                  </font></td>
              </tr>
              <tr> 
                <td class=C width="23%" bgcolor=#f3f3f3>地址：</td>
                <td class=C width="77%" 
                  bgcolor=#f3f3f3><%=rs_q("address")%></td>
              </tr>
              <tr> 
                <td class=C width="23%">邮政编码：</td>
                <td class=C width="77%"><%=rs_q("post")%></td>
              </tr>
              <tr> 
                <td class=C width="23%">网址：</td>
                <td class=C 
width="77%"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
      </table>
      <h3>&nbsp;</h3>
    </td>
  </tr>
</table>
        


<%rs_q.close 
set rs_q=nothing
conn.close
set conn=nothing
%>
<%end if%>
<!--#include file="Copyright.htm"--></BODY></HTML>
