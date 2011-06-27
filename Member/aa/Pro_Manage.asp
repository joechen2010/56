<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<title>管理产品信息</title>
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
<BODY text=#000000 leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" bgcolor="#F7F7FF">
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
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 管理产品信息</td>
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
                        <div align="center"><font color="#FFFFFF">管理产品信息</font></div>
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
sfilename="Pro_Manage.asp"
strUnit="个产品"
const MaxPerPage=3
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
dim gsrs
dim gssql
dim qymc
gssql="select * from qyml where id="+cstr(gsid)+" order by ID desc"
Set gsrs= Server.CreateObject("ADODB.Recordset") 
gsrs.open gssql,conn,1,1 
qymc=gsrs("qymc")
gsrs.close
sql="select * from spzs where gsid="+cstr(gsid)+" order by ID desc"
set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
if rs.eof and rs.bof then
    response.write "<p>&nbsp;</p><p align=center>您暂时还没有发布信息！<p>"
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
call showpage(sfilename,totalPut,maxperpage,true,true,strUnit)
else 
if (currentPage-1)*MaxPerPage<totalPut then 
rs.move (currentPage-1)*MaxPerPage 
dim bookmark 
bookmark=rs.bookmark 
showContent 
call showpage(sfilename,totalPut,maxperpage,true,true,strUnit)
else 
currentPage=1 
showContent 
call showpage(sfilename,totalPut,maxperpage,true,true,strUnit)
end if 
end if 
end if
conn.close
set conn=nothing 
sub showContent
%>
                  <table border="0" cellspacing="0" cellpadding="0" width=100% height="22" style="border-collapse: collapse" bordercolor="#111111">
                    <tr align=center height=30 bgcolor=#ececec> 
    <td height="22" width="600" bgcolor="#CEDBFF">
<p><font class="f12" color="#ff6600"><span style="font-size: 10.5pt"><b><%=session("qymc")%>的产品管理</b></span></font></td></tr> 
</table>
                  <table border="0" cellspacing="0" width="100%" style="border-collapse: collapse" height="1">
                    <tr height=30><td width="180" height="12"></td></tr> <%do while not rs.eof%> <tr> 
<td width="180" height="1" rowspan="3" bgcolor="#F7F7FF"> 
                      <p align="center"> <a target="_blank" href="../../Product/show_product.asp?id=<%=rs("id")%>"> 
                        <IMG 
border=0 
src="../../picture/<%=rs("picture")%>" width="125" height="108" alt="<%=rs("cpmc")%>" style="border: 1px solid #000000"></a>
                    </td><td width="428" height="22" colspan="4" bgcolor="#F7F7FF">　</tr> 
<tr> <td width="428" height="96" colspan="4" valign="bottom" bgcolor="#F7F7FF" > 
<div align="left"> <table border="1" style="border-collapse: collapse" bordercolor="#111111" width="91%" id="AutoNumber4" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="3" height="80"> 
<tr> 
                            <td width="56%" height="17" style="color: #000; font-family: 宋体; font-size: 10.5pt; line-height:100%" bgcolor="#F0F0FF">&nbsp;<span style="font-size: 10.5pt"> 
                              <a target="_blank" href="../../Product/show_product.asp?id=<%=rs("id")%>"><%=rs("cpmc")%></a></span></td>
                            <td width="22%" height="17" style="color: #000; font-family: line-height:100%" bgcolor="#F0F0FF"> 
<p align="center"><%=rs("idate")%></td></tr> <tr> <td width="78%" height="17" style="color: #000; font-family: 宋体; font-size: 12px; line-height:100%" bgcolor="#F0F0FF" colspan="2" >&nbsp; 
<%=mid(rs("jysm"),1,22)%>...</td></tr> <tr> <td width="78%" height="17" style="color: #000; font-family: 宋体; font-size: 12px; line-height:100%" bgcolor="#F0F0FF" colspan="2">&nbsp; 
交易价格：<font color="#FF6600"><%if rs("cpjg")="0" then response.write"价优" else response.write""&rs("cpjg")&"元" end if%></font></td></tr> 
<tr> <td width="78%" height="17" style="color: #000; font-family: 宋体; font-size: 12px; line-height:100%" bgcolor="#F0F0FF" colspan="2"> 
&nbsp; 产品产地：<%=rs("cpcd")%></td></tr> <tr> <td width="78%" height="1" style="color: #000; font-family: 宋体; font-size: 12px; line-height:100%" bgcolor="#F0F0FF" colspan="2"> 
&nbsp; <a target="_blank" href="../../showroom/index.asp?id=<%=session("id")%>"><%=rs("sccj")%></a></td></tr> 
</table></div></td></tr> <tr> <td width="308" height="1" bgcolor="#F7F7FF"></td>
                    <td width="53" height="1" bgcolor="#F7F7FF"> <a href="Pro_edit.asp?id=<%=rs("id")%>"> 
                      <img border="0" src="images/Edit_2.gif"></a></td>
                    <td width="52" height="1" bgcolor="#F7F7FF"><a href="Pro_del.asp?id=<%=rs("id")%>"><img border="0" src="images/A_delete.gif"></a></td>
                    <td width="19" height="1" valign="middle" bgcolor="#F7F7FF"> 
<p align="center"></td></tr> <tr> <td width="100%" colspan="5" bgcolor="#FFFFFF"> 
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5"> 
<tr> <TD height="1" width="100%" bgcolor="#F4FAFF" style="font-family: 宋体; font-size: 9pt; line-height: 14pt"> 
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber7"> 
<tr> <TD background=../../images/bg_b.gif style="color: #000; font-family: 宋体; font-size: 12px; line-height:100%" bgcolor="#EEF7FF"><IMG 
height=1 src="" width=1></TD></tr> </table></TD></tr> </table></td></tr> <% i=i+1 
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
rs.close 
set rs=nothing 
%> </table><%
end sub 
%>
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