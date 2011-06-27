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
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
</HEAD>
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
          <td background="images/skin3_15.jpg" width="167" height="27" align="center"><a class=white 
                        href="aboutus.asp?id=<%=gsid%>"><font 
                        color=#ffffff><b>公司简介</b></font></a></td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="96%" border="0" cellspacing="1" cellpadding="0" bgcolor="239687">
        <tr>
          <td bgcolor="239687" height="3"></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">
            <TABLE class=style cellSpacing=5 cellPadding=0 width="100%" 
                  align=center border=0>
              <TBODY> 
              <TR>
                      <TD class=S2 align=middle bgColor=#f9f9f9></TD></TR>
                    <TR>
                      <TD class=S2 bgColor=#f9f9f9 style="width=800;WORD-WRAP: break-word;word-break: break-all">
                        <%if rs_q("qyjj")<>"" then 
        response.write Replace(rs_q("qyjj"),chr(13),"<br>&nbsp;&nbsp;&nbsp;")
	else
	response.write rs_q("qyjj")
end if
%></TD></TR>
                    <TR>
                      <TD class=S>
                        <TABLE height=100 cellSpacing=1 cellPadding=1 
                        width="100%" align=center bgColor=#ffcccc border=0>
                          <TBODY>
                          <TR>
                            <TD bgColor=#ffffff>
                              <TABLE cellSpacing=0 cellPadding=5 width="100%" 
                              border=0>
                                <TBODY>
                                <TR>
                                <TD bgColor=#feedd1 height=100>
                                <TABLE class=style cellSpacing=5 cellPadding=5 
                                width="100%" align=center bgColor=#ffffff 
                                border=0>
                                <TBODY>
                                <TR>
                                        <TD height=85> 
                                          <table cellspacing=2 cellpadding=3 width="100%" align=center border=0>
                                            <tbody> 
                                            <tr> 
                                              <td valign=top align=left width="50%" bgcolor=#fff4e5>主营产品或服务： 
                                                <br>
                                                <span class=other><%=rs_q("Zycp")%></span></td>
                                              <td valign=top align=left width="50%" bgcolor=#fff4e5>企业经营性质： 
                                                <br>
                                                <span class=other><%=rs_q("qyxz")%> 
                                                </span></td>
                                            </tr>
                                            <tr> 
                                              <td background="" colspan=2 height=1></td>
                                            </tr>
                                            <tr> 
                                              <td valign=top align=left width="50%">经营模式： 
                                                <br>
                                                <span 
            class=other><%=rs_q("Qylb")%></span></td>
                                              <td valign=top align=left>法人代表/企业负责人： 
                                                <br>
                                                <span class=other> <%=rs_q("Frdb")%></span></td>
                                            </tr>
                                            <tr> 
                                              <td background="" colspan=2 height=1></td>
                                            </tr>
                                            <tr> 
                                              <td valign=top align=left width="50%" bgcolor=#fff4e5>公司注册地： 
                                                <br>
                                          <span class=other><%=rs_q("province")%><%=rs_q("city")%><%=rs_q("Area")%> 
                                          </span></td>
                                              <td valign=top align=left bgcolor=#fff4e5>公司经营地： 
                                                <br>
                                                <span 
            class=other><%=rs_q("address")%> </span></td>
                                            </tr>
                                            <tr> 
                                              <td background="" colspan=2 height=1></td>
                                            </tr>
                                            <tr> 
                                              <td valign=top width="50%">员工人数： 
                                                <br>
                                                <span class=other><%=rs_q("Ygrs")%></span></td>
                                              <td valign=top align=left>注册资金： 
                                                <br>
                                                <span class=other> <%=rs_q("Zczj")%> 
                                                万元</span></td>
                                            </tr>
                                            <tr> 
                                              <td background="" colspan=2 height=1></td>
                                            </tr>
                                            </tbody> 
                                          </table>
                                        </TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></td>
        </tr>
      </table>
      <h3>&nbsp;</h3>
    </td>
  </tr>
</table>
        

<!--#include file="Copyright.htm"-->
<%end if%>