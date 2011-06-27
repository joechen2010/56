<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../Member/check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>信息中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="refresh" content="60;URL=infocenter_main.asp">
<style type="text/css">
.navPoint {COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt}
</style>
</head>
<body scroll=no topmargin=0  rightmargin=0  leftmargin=0>
<%if request("xs")="" then%>
<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
  <tr> 
    <td>
	<iframe frameborder="0" id="main" name="main" scrolling="yes" src="infocenter_goods.asp" style="HEIGHT: 90%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"> 
      </iframe></td>
  </tr>
  <tr> 
    <td>
	<iframe frameborder="0" id="main" name="main" scrolling="yes" src="infocenter_car.asp" style="HEIGHT: 90%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"> 
      </iframe>
	  </td>
  </tr>
 </table>
<%end if%>
<%if request("xs")="goods" then%>
<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
  <tr> 
    <td>
	<iframe frameborder="0" id="main" name="main" scrolling="yes" src="infocenter_goods.asp" style="HEIGHT: 90%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"> 
      </iframe></td>
  </tr>
 </table>
<%end if%>
<%if request("xs")="car" then%>
<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
  <tr> 
    <td>
	<iframe frameborder="0" id="main" name="main" scrolling="yes" src="infocenter_car.asp" style="HEIGHT: 90%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"> 
      </iframe>
	  </td>
  </tr>
 </table>
<%end if%>
</body>
</html>
