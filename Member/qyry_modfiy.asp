<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
 <%session("id")=request("id")
sql="select * from qyml where user='"+session("user")+"' order by ID desc"
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 

sqls="select * from qyzs where id="+request("id")+""
Set rss= Server.CreateObject("ADODB.Recordset") 
rss.open sqls,conn,1,1 


%>

<html>
<head>
<TITLE>诚信物流网 会员控制中心 - 物流商人的网上家园</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<style type="text/css">
<!--
body {
	background-color: #2C68B1;
	margin: 0px;
}
-->
</style>
<link href="images/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
<script language="javascript" src="images/ShowProcessBar.js"></script> 
<script language=javascript>
function checkImage(sId)
{
  if(( document.all[sId].value.indexOf(".gif") == -1) && (document.all[sId].value.indexOf(".jpg") == -1)) {
    alert("请选择gif或jpg的图象文件");
    event.returnValue = false;
    }
}
</script>
</head>
<BODY>
<!--#include file="head.htm"-->
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 上传企业证书</td>
          <td width="17%"><img src="images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="534"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> <%=session("user")%> </font> </strong>，欢迎您回来！</td>
        </tr>
      </table>
      <table width="100%" height="497"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top"><table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="images/title_2.jpg" bgcolor="#FFFFFF" >
                        <div align="center"><font color="#FFFFFF">上传企业证书</font></div>
                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
                <td valign="top" bgcolor="#FFFFFF"> 
                  <table border="0" cellspacing="1" style="border-collapse: collapse" width="100%" id="AutoNumber2" bordercolorlight="#FFEEB3" cellpadding="5" bordercolordark="#FFEEB3" bgcolor="#B1D4F2">
                    <tr bgcolor="#EDF6FF"> 
                      <TD height=32 width=135 align="center">上传图片:</TD>
                      <td width="443" height="32" bgcolor="#EDF6FF"> 
                        <form name="form" method="post" action="qyryupfilemodfiy.asp" >
                          <p> 
                            <input type="hidden" name="filepath" value="ryzs/">
                            <input type="hidden" name="act" value="upload">
                            证书名称： 
                            <input name="zsmc" type="text" value="<%=rss("zsmc")%>">
                          </p>
                          <p>
                            <iframe name="smallimg" frameborder=0 width=450 height=20 scrolling=no src=smallupload.asp></iframe>
                            <br>
              </p>
                          <p>图片路径： 
                            <input name="smallimg" type="TEXT" id="smallimg"  style="background-color:ffffff;color:000000;border: 1 double" value="<%=rss("zsurl")%>" size=34 maxlength=100 readonly>
                          </p>
                          <p align="center">
                            <input type="submit" name="Submit" value="修　改">
                          </p>
                        </form>                      </td>
                    </tr>
                  </table>
                </td>
            </tr>
          </table>
                        
          </td>
        </tr>
      </table>
</td>
  </tr>
</table>

<!--#include file="bottom.htm"-->
</body>
</html>