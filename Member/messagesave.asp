<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="check.asp"-->
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
<TITLE>诚信物流网 会员控制中心 - 物流商人的网上家园</TITLE>
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
</head>
<body>
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
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 修改公司资料</td>
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
                        <div align="center"><font color="#FFFFFF">修改公司资料</font></div>
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
			  <table width="100%"  border="0">
                             <tr>
                              <td align="left"><div id="TipAll" >
                               请输入所有标有*的项目。
                              </div></td>
                            </tr>
                           </table>
                  <p align=center>信息发送成功！</p>
                  <p> 还不是诚信物流网的会员？现在就加入，享受更多会员服务！ * 无限制访问产品及供应商的数据库 * 我的办公室轻松管理在线贸易、收发信息、修改资料 
                    * 直接联系商家，发布商业信息，订阅产品速递 * 了解国际贸易法律法规 * 获得海内外贸易发展状况 还不是诚信物流网的会员？请点此注册，保存您在诚信物流网上将要收发的所有信息。 
                    已经是诚信物流网的会员？请点此登录，保存您在诚信物流网上将要收发的所有信息。 </p>
                  <p align=center><a href="javascript:window.history.back()">返回</a></p>
</td>
            </tr>
          </table>
                        <table width="534"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="right"><a href="#top"><img src="images/icon_top.gif" alt="回到页面顶部" border="0" align="absmiddle" width="15" height="18"></a> 
                  <a href="#top">再检查一遍</a></td>
              </tr>
            </table></td>
        </tr>
      </table>
</td>
  </tr>
</table>

<!--#include file="bottom.htm"-->
</body>
</html>
