<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../Member/check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>诚信物流网-信息中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
.navPoint {COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt}
</style>
<script>
function switchSysBar(){
if (switchPoint.innerText==3){
switchPoint.innerText=4
document.all("frmTitle").style.display="none"
}else{
switchPoint.innerText=3
document.all("frmTitle").style.display=""
}}
</script>
</head>

<body scroll=no topmargin=0  rightmargin=0  leftmargin=0>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20%"><img src="../images/logo.gif" border=0></td>
	<td width="15%">你好：<%=session("name")%></td>
	<td width="8%"><a href="infocenter_index.asp?left=">即时配货</a></td>
    <td width="8%"><a href="infocenter_index.asp?left=tj">推荐配货</a></td>
    <td width="8%"><a href="infocenter_index.asp?left=fb">发布配货</a></td>
    <td width="9%"><a href="infocenter_index.asp?left=search">查询配货</a></td>
    <td width="8%"><a href="infocenter_index.asp?left=gq">发布供求</a></td>
    <td width="14%"><a href="infocenter_index.asp?left=searchgq">查询供求</a></td>
    <td width="10%">&nbsp;</td>
  </tr>    
</table>

<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
  <tr> 
    <td align="middle" id="frmTitle" noWrap valign="center" name="frmTitle">
	<%if request("left")="" then%>
	<iframe frameborder="0" id="carnoc" name="carnoc" scrolling="no"  src="infocenter_left.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2">      </iframe>
	  <%end if%>
	  <%if request("left")="tj" then%>
	  <iframe frameborder="0" id="carnoc" name="carnoc" scrolling="no"  src="infocenter_left_tj.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2"> </iframe>
	  <%end if%>
	  <%if request("left")="fb" then%>
	<iframe frameborder="0" id="carnoc" name="carnoc" scrolling="no"  src="infocenter_left_fb.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 200px; Z-INDEX: 2">      </iframe>	
	  <%end if%>
	  <%if request("left")="search" then%>
	<iframe frameborder="0" id="carnoc" name="carnoc" scrolling="no"  src="infocenter_left_search.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 300px; Z-INDEX: 2">      </iframe>	  
	  <%end if%>
	  <%if request("left")="searchgq" then%>
	<iframe frameborder="0" id="carnoc" name="carnoc" scrolling="no"  src="infocenter_left_searchgq.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 300px; Z-INDEX: 2">      </iframe>	  
	  <%end if%>
	  <%if request("left")="gq" then%>
	<iframe frameborder="0" id="carnoc" name="carnoc" scrolling="no"  src="infocenter_left_gq.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 200px; Z-INDEX: 2">      </iframe>	  
	  <%end if%>    </td>
    <td style="WIDTH: 9pt" bgcolor="#A4B6D7"> 
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr> 
          <td style="HEIGHT: 100%" onClick="switchSysBar()"> <font style="FONT-SIZE: 9pt; CURSOR: default; COLOR: #ffffff"> 
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <span class="navPoint" id="switchPoint" title="关闭/打开左栏">3</span><br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            屏幕切换 </font></td>
        </tr>
      </table>    </td>
    <td style="WIDTH: 100%">
	<%if request("left")="" then%>
	<iframe frameborder="0" id="main" name="main" scrolling="yes" src="infocenter_main.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">      </iframe>
	  <%end if%>
	<%if request("left")="tj" then%>
	<iframe frameborder="0" id="main" name="main" scrolling="yes" src="infocenter_pp.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">      </iframe>
	  <%end if%>	 
	<%if request("left")="search" then%>
	<iframe frameborder="0" id="main" name="main" scrolling="yes" src="infocenter_search.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">      </iframe>
	  <%end if%>
	  <%if request("left")="fb" then%>
	<iframe frameborder="0" id="main" name="main" scrolling="yes" src="car_manage.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">      </iframe>
	  <%end if%>
	  <%if request("left")="gq" then%>
	<iframe frameborder="0" id="main" name="main" scrolling="yes" src="gq_manage.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">      </iframe>
	  <%end if%>	  	  
	  <%if request("left")="searchgq" then%>
	<iframe frameborder="0" id="main" name="main" scrolling="yes" src="infocenter_gq.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">      </iframe>
	  <%end if%>	</td>
  </tr>
</table>
</body>
</html>
