<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<%
dim subject,content,F_Company,F_Name,F_Email,F_address,F_Post,F_Tel,F_Fax,T_User,F_User,TF_Time
subject=request("subject")
content=request("content")
F_Company=request("F_Company")
F_Name=request("F_Name")
F_Email=request("F_Email")
F_address=request("F_address")
F_Post=request("F_Post")
F_Tel=request("F_Tel")
F_Fax=request("F_Fax")
T_User=request("T_User")
F_User=request("F_User")
T_Name=request("T_Name")
T_Company=request("T_Company")
T_Email=request("T_Email")
set rs=server.createobject("adodb.recordset")
sql="select * from message"
rs.open sql,conn,1,3
rs.addnew
rs("subject")=subject
rs("content")=content
rs("F_Company")=F_Company
rs("F_Name")=F_Name
rs("F_Email")=F_Email
rs("F_address")=F_address
rs("F_Post")=F_Post
rs("F_Tel")=F_Tel
rs("F_Fax")=F_Fax
rs("T_User")=T_User
rs("F_User")=F_User
rs("T_Name")=T_Name
rs("T_Email")=T_Email
rs("T_Company")=T_Company
rs("TF_Time")=now()
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>信息发送成功！</title>
</head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/mian.css" rel=stylesheet type=text/css> 
<STYLE type=text/css>.p14 {
	FONT-SIZE: 14px; COLOR: #000000; LINE-HEIGHT: 130%
}
</STYLE>
<body leftmargin="0" topmargin="0">
<!--#include file="../inc/head1.htm"-->
  <TABLE cellSpacing=0 cellPadding=0 width=770 align=center border=0>
  <TBODY>
  <TR>
      <TD class=S height=30>您当前位置：<A href="../">首页</A> &gt;&gt; 我要留言 &gt;&gt; 
        留言成功</TD>
    <TD>&nbsp;</TD>
    <TD>&nbsp;</TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=770 align=center 
background=images/inquiry_t_bg.gif border=0>
  <TBODY>
  <TR>
    <TD><IMG height=60 src="images/inquiry_t_l.gif" border=0></TD>
    <TD width=13><IMG height=60 src="images/inquiry_t_r.gif" 
      width=13></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=770 align=center border=0>
  <TBODY>
  <TR>
    <TD width=5 background=images/pro_r_l_bg.gif><IMG height=1 
      src="images/spacer.gif" width=5></TD>
    <TD align=middle height=200>
      <TABLE class=table3 cellSpacing=0 cellPadding=0 width="94%" align=center 
      border=0>
        <TBODY>
        <TR>
          <TD width="25%" height=180>
              <DIV align=center><IMG height=128 src="images/sendmsg.jpg" 
            width=128></DIV>
            </TD>
          <TD width="75%">
            <TABLE cellSpacing=0 cellPadding=0 width="90%" border=0>
              <TBODY>
              <TR>
                <TD class=buttoncolor>您的询价信息已经成功发出！请耐心等待回复！</TD></TR>
              <TR>
                <TD>&nbsp;</TD></TR>
              <TR>
                <TD height=30>
                    <DIV align=center><A 
                  href="../Trade/Default.asp"><STRONG>查看诚信物流网其它供求信息！</STRONG></A></DIV>
                  </TD></TR>
              <TR>
                <TD height=30>
                  <DIV align=center><A 
                  href="javascript:history.go(-2)"><STRONG>返回上一层继续查看信息</STRONG></A></DIV></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
    <TD width=6 background=images/pro_r_r_bg.gif><IMG height=8 
      src="images/spacer.gif" width=6></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=770 align=center 
background=images/pro_r_b_bg.gif border=0>
  <TBODY>
  <TR>
    <TD width=11><IMG height=17 src="images/pro_r_b_l.gif" 
    width=11></TD>
    <TD><IMG height=1 src="images/spacer.gif" width=1></TD>
    <TD width=13><IMG height=17 src="images/pro_r_b_r.gif" 
    width=13></TD></TR>
	</TBODY>
	</TABLE>
<!--#include file="../inc/down.htm"--> 
</body>
</html>