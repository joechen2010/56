<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>后台管理首页</title>
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
<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
  <tr> 
    <td align="middle" id="frmTitle" noWrap valign="center" name="frmTitle"> <iframe frameborder="0" id="carnoc" name="carnoc" scrolling="yes"  src="admin_left.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2"> 
      </iframe> </td>
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
      </table>
    </td>
    <td style="WIDTH: 100%"> <iframe frameborder="0" id="main" name="main" scrolling="yes" src="admin_main.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"> 
      </iframe></td>
  </tr>
</table>
</body>
</html>
