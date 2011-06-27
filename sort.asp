<%@ codepage ="936" %>
<!--#include file="inc/conn2.asp"-->
<!--#include file="inc/function.asp"-->
<%
'conn.execute("update message set new=1 where id="&request("id"))
set rs=server.createobject("adodb.recordset")
'sql="select a.gsid,count(*) as num from info2 a,qyml b,zx_info c where a.gsid =b.user and  b.user=c.gsid group by a.gsid "
sql="select top 10 gsid,count(gsid) from (select gsid from info2 union all select gsid from zp_info union all select gsid from zx_info union all select gsid from repair_info union all select gsid from file_info union all select gsid from gq_info union all select gsid from gq_info) as total  group by gsid order by count(gsid) desc"
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
<table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93" align="center">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
                  <table cellspacing=1 cellpadding=3 width=100% align=center border=0 bgcolor="#B1D4F2">
                    <tr bgcolor="#EDF6FF">
                      <td align="center">名次</td>
                      <td align="center">会员帐号</td>
                      <td align="center">会员公司</td>
                      <td align="center">发布信息条数</td>
                    </tr>
					<%i=1%>
					<%do while not rs.eof%>
                    <tr bgcolor="#EDF6FF"> 
                     <td width="453" align="center"><%=i%></td>
                     <td width="453" align="center"><%=rs(0)%></td>
                     <td width="453" align="center">
					 <%set rs2=server.CreateObject("adodb.recordset")
					   sql2="select * from qyml where user='"&rs(0)&"'"
					   rs2.open sql2,conn,1,1
					   if not (rs2.eof and rs2.bof) then
					 %>
				  <%if len(trim(rs2("qymc")))>10 then%>
				  <a href="yellowpage/site.asp?InfoID=<%=rs2("ID")%>&user=<%=rs2("user")%>" target="_blank"><%=left(rs2("qymc"),8)%>...</a>
				  <%else%>
				  <a href="yellowpage/site.asp?InfoID=<%=rs2("ID")%>&user=<%=rs2("user")%>" target="_blank"><%=rs2("qymc")%></a>
				  <%end if%>
				  <%end if%>					 
					 </td>
                     <td width="453" align="center"><%=rs(1)%></td>
                    </tr> 				  
                      <%
					  i=i+1
					  rs.movenext
					  loop
					  %>                  
                    <tr bgcolor="#EDF6FF"> 
                      <td colspan="4" align="center"> <a href="javascript:window.close()">关闭窗口</a></td>
                      </tr>
				</table>				
				</td>
            </tr>
          </table>

</body>
</html>
<%end if%>
