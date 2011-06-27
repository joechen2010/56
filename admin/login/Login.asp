<%@language=vbscript codepage=936 %>
<%
option explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<HTML>
<HEAD>
<TITLE>登录系统</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>.font1 {
	FONT-SIZE: 12px; LINE-HEIGHT: 130%; FONT-FAMILY: "宋体"
}
A {
	FONT-SIZE: 12px; FONT-FAMILY: "宋体"
}
A:link {
	FONT-SIZE: 12px; COLOR: #6262b1; TEXT-DECORATION: none
}
A:visited {
	FONT-SIZE: 12px; COLOR: #6262b1; TEXT-DECORATION: none
}
A:hover {
	FONT-SIZE: 12px; COLOR: #ffffff; TEXT-DECORATION: underline
}
.input {
	BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BORDER-LEFT: #000000 1px solid; COLOR: #000000; BORDER-BOTTOM: #000000 1px solid; FONT-FAMILY: "宋体"; BACKGROUND-COLOR: #ffffff
}
</STYLE>
<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.UserName.value=="")
	document.Login.UserName.focus();
else
	document.Login.UserName.select();
}
function CheckForm()
{
	if(document.Login.UserName.value=="")
	{
		alert("请输入用户名！");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("请输入密码！");
		document.Login.Password.focus();
		return false;
	}
}
//-->
</script>
<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY text=#000000 bgColor=#002779 leftMargin=0 topMargin=0 marginheight="0" 
marginwidth="0">
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD vAlign=center align=center> 
      <TABLE cellSpacing=0 cellPadding=0 width=460 border=0>
        <TBODY>
        <TR>
          <TD><IMG height=23 src="images/login_1.jpg" 
        width=190></TD>
        </TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=460 border=0>
        <TBODY>
        <TR>
          <TD><IMG height=142 src="images/login_2.jpg" 
        width=460></TD>
        </TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=460 border=0>
        <TBODY>
        <TR>
          <TD bgColor=#eeeeee height=6></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=460 border=0>
        <TBODY>
        <TR>
          <TD align=middle bgColor=#ffffff height=120>
            <TABLE cellSpacing=0 cellPadding=0 width=300 align=center 
              border=0>
<form name="Login" action="Chklogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
              <TR>
                <TD align=middle colSpan=3 height=5></TD></TR>
              <TR>
                <TD class=nwes width=5 height=36></TD>
                <TD class=font1 width=56 height=36>用户名</TD>
                <TD>
                    <INPUT class=input maxLength=12 name=UserName>
                     </TD></TR>
              <TR>
                <TD class=nwes height=36>&nbsp; </TD>
                <TD class=font1 height=36>口　令</TD>
                <TD>
                    <INPUT class=input type=password maxLength=12 
                  name=Password>
                     </TD></TR>
              <TR>
                <TD colSpan=3 height=5></TD></TR>
              <TR>
                <TD>&nbsp; </TD>
                <TD align=middle>&nbsp; </TD>
                <TD>
                    <INPUT type=image height=16 width=70 
                  src="images/bt_login.gif" border=0 name=imageField> 
              </TD></TR></FORM></TABLE></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=460 bgColor=#ffffff border=0>
        <TBODY>
        <TR>
          <TD><IMG height=10 src="images/login_3.gif" width=10></TD>
          <TD align=right><IMG height=10 src="images/login_4.gif" 
          width=10></TD>
        </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
</BODY></HTML>
 
