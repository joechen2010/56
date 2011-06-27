<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<title>已发布信息</title>
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
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"> 
<!--#include file="head.htm"-->
<table width="750"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 已发布信息</td>
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
                        <div align="center"><font color="#FFFFFF">已发布信息</font></div>
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
<%
const MaxPerPage=5
dim totalPut 
dim CurrentPage
dim TotalPages
dim i
if not isempty(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if
dim sql
dim rs
dim qymc
sql="select * from info where gsid="+cstr(gsid)+" order by info_ID desc"
set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
if rs.eof and rs.bof then 
    response.write"<p>&nbsp;</P><p align=center>暂无信息发布！</p>"
response.end 
else 
totalPut=rs.recordcount 
if currentpage<1 then 
currentpage=1 
end if 
if (currentpage-1)*MaxPerPage>totalput then 
if (totalPut mod MaxPerPage)=0 then 
currentpage= totalPut \ MaxPerPage 
else 
currentpage= totalPut \ MaxPerPage + 1 
end if 
end if 
if currentPage=1 then 
showContent 
showpage totalput,MaxPerPage,"info_manage.asp"

else 
if (currentPage-1)*MaxPerPage<totalPut then 
rs.move (currentPage-1)*MaxPerPage 
dim bookmark 
bookmark=rs.bookmark 
showContent 
showpage totalput,MaxPerPage,"info_manage.asp"
else 
currentPage=1 
showContent 
showpage totalput,MaxPerPage,"info_manage.asp"
end if 
end if 
end if 
sub showContent 
%>
                  <table border="0" cellspacing="1" style="border-collapse: collapse" width="100%" id="AutoNumber6" cellpadding="5">
                    <tr> <td align="center" width="100%"> <div align="center"> <center>
                          </center></div>
                        <%do while not rs.eof%>
                        <TABLE border=0 cellSpacing=1 width="100%" style="border-collapse: collapse" height="75" cellpadding="5">
                          <TBODY> 
                          <TR height="30"> 
                            <TD align=left height="21" width="782" colspan="3" bgcolor="#F4FAFF"> 
                              <img border="0" src="images/logo_min.gif"><b><span style="font-size: 10.5pt"><a href="../trade/business_info.asp?info_id=<%=rs("info_id")%>"><%=rs("showname")%></a></span></b><span style="font-size: 10.5pt"> 
                              </span><%=rs("dateandtime")%><font color="#fc9603">&nbsp;&nbsp; 
                              【<%=rs("type")%>】</font></TD>
                          </TR>
                          <TR height="30"> 
                            <TD height="25" width="782" colspan="3" bgcolor="#E4F4FC"> 
                              <%=mid(rs("content"),1,88)%>..<a href="../trade/business_info.asp?info_id=<%=rs("info_id")%>">[详细信息]</a> 
                            </TD>
                          </TR>
                           <tr> <TD height="17" width="669" bgcolor="#F4FAFF"></TD>
              <TD height="17" width="57" bgcolor="#F4FAFF"> <a href="info_edit.asp?info_id=<%=rs("info_id")%>"> 
                <img border="0" src="images/Edit_2.gif" width="47" height="18"></a></TD>
              <TD height="17" width="56" bgcolor="#F4FAFF"> <a href="info_del.ASP?ID=<%=rs("info_id")%>"> 
                <img border="0" src="images/A_delete.gif" width="52" height="16"></a></TD>
            </tr> 
<tr> 
                            <TD height="11" width="782" bgcolor="#F4FAFF" colspan="3">&nbsp; 
                            </TD>
                          </tr> </TBODY> </TABLE><% i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
rs.close 
set rs=nothing 
%> <%
end sub 
function showpage(totalnumber,maxperpage,filename) 
dim n 
if totalnumber mod maxperpage=0 then 
n= totalnumber \ maxperpage 
else 
n= totalnumber \ maxperpage+1 
end if 
response.write "<form method=Post action="&filename&">" 
if CurrentPage<2 then 
response.write "<table border=0 width=95% cellspacing=0 cellpadding=0>"
response.write "<tr height=5>"
response.write "<td align=left valign=top>"
response.write "<font color=red><b>共"&n&"页&nbsp;第"&CurrentPage&"页 " 
response.write "共发布"&totalnumber&"&nbsp;条信息</b></font>" 
response.write "</td>"
response.write "<td VALIGN=middle align=right >"
response.write "【最前页】【上一页】" 
else 
response.write "<table border=0 width=95% cellspacing=0 cellpadding=0>"
response.write "<tr height=5>"
response.write "<td align=left valign=top>"
response.write "<font color=red><b>共"&n&"页&nbsp;第"&CurrentPage&"页 " 
response.write "共发布"&totalnumber&"&nbsp条信息</b></font>" 
response.write "</td>"
response.write "<td VALIGN=center align=right >"
response.write "【<a href="&filename&"?page=1&id="&gsid&">最前页</a>】" 
response.write "【<a href="&filename&"?page="&CurrentPage-1&"&id="&gsid&">上一页</a>】" 
end if 
if n-currentpage<1 then 
response.write "【下一页】【最后页】" 
else 
response.write "【<a href="&filename&"?page="&(CurrentPage+1)&"&id="&gsid&">" 
response.write "下一页</a>】【 <a href="&filename&"?page="&n&"&id="&gsid&">最后页</a>】" 
end if 
response.write "</form>" 
response.write "</td>"
response.write "</tr>"
response.write "</table>"
end function
%> </tr> </table>
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

<!--#include file="bottom.htm"--></body></html>