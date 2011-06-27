<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../Member/check.asp"-->
<%
conn.execute("update message set new=1 where id="&request("id"))
set rs=server.createobject("adodb.recordset")
sql="select * from message where id="&request("id")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
    response.write "没有数据，请<a href=""javascript:windows.history.back()"">返回</a>"
else
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
<link href="../Member/images/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="566" align="center" valign="top" bgcolor="#FFFFFF" id="main">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="../Member/images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 发出的留言信息</td>
          <td width="17%"><img src="../Member/images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="100%" height="243"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top"><table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="../Member/images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="../Member/images/title_2.jpg" bgcolor="#FFFFFF" >
                        <div align="center"><font color="#FFFFFF"> 发出的留言信息</font></div>                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="../Member/images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
                  <table cellspacing=1 cellpadding=3 width=100% align=center border=0 bgcolor="#B1D4F2">
                    <tr bgcolor="#EDF6FF"> 
<%set rs2=server.createobject("adodb.recordset")
sql2="select * from qyml where user='"&rs("F_user")&"'"		
rs2.open sql2,conn,1,1	
if rs2.eof and rs2.bof then
    response.write "没有数据"
else		  
%>                       <td colspan="2" align="center"><font color="#FF0000"><%=rs2("qymc")%></font>给我的留言</td>
<%end if%>					  
                    </tr>                    
					<tr bordercolor=#ffffcc bgcolor="#EDF6FF"> 
                      <td width=54><b>主题:</b></td>
                      <td width="453"><b><%=rs("subject")%></b></td>
                    </tr>
                    <tr valign=top bgcolor="#EDF6FF"> 
                      <td valign=top width=54 height=120>内容：<br>    </td>
                      <td height=120><%=dvHTMLEncode(rs("content"))%></td>
                    </tr>
                    <tr bgcolor="#EDF6FF"> 
                      <td colspan="2" align="center"> 
<a href="../Member/reply.asp?id=<%=rs("id")%>"></a>
<a href="javascript:window.close()">返回</a></td>
                    </tr>
				</table>				
				</td>
            </tr>
          </table>
          </td>
        </tr>
      </table></td>
  </tr>
</table>

</body>
</html>
<%end if%>
