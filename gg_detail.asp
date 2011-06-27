<%@ codepage ="936" %>
<!--#include file="inc/conn2.asp"-->
<!--#include file="inc/function.asp"-->
<%
'conn.execute("update message set new=1 where id="&request("id"))
set rs=server.createobject("adodb.recordset")
sql="select * from gonggao where id="&request("id")
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
<link href="Member/images/main.css" rel="stylesheet" type="text/css">
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
      <table width="100%" height="428"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top"><table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
                  <table cellspacing=1 cellpadding=3 width=100% align=center border=0 bgcolor="#B1D4F2">
                    <tr bgcolor="#EDF6FF"> 
                     <td width="453" colspan="2" align="center"><%=rs("bt")%></td>
                    </tr> 				  
                    <tr bgcolor="#EDF6FF"> 
                     <td width="453" colspan="2" align="center"><%=rs("gonggao")%></td>
                    </tr>                    
                    <tr bgcolor="#EDF6FF"> 
                      <td colspan="2" align="center"> <a href="javascript:window.close()">关闭窗口</a></td>
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
