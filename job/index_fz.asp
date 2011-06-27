
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
fz=request("fz")
%>
<HTML>
<HEAD>
<title>专线信息检索条件</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<STYLE type=text/css>.bg {
	BACKGROUND-POSITION: 50% top; BACKGROUND-REPEAT: no-repeat
}
</STYLE>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }

</script>
</HEAD>
<body leftmargin="0" rightmargin="0" style="FONT-SIZE: 10pt" bgColor="#FFFFFF" topmargin="0">
<!--#include file="../inc/top_fz.html"-->
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td valign="top" bgcolor="#FFFFBD"><iframe name="left" src="left.asp" width="100%" frameborder="0" scrolling="no"></iframe></td>
    <td width="500"><iframe name="right" src="right_fz.asp" width="100%" frameborder="0" scrolling="no" height="580"></iframe></td>
  </tr>
</table>
<!--#include file="../inc/down.htm"-->
</body>
</HTML>
