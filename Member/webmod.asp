<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<%
dim totalPut 
dim CurrentPage
dim TotalPages
dim i,j
MaxPerPage=9
if not isempty(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if
%> 
<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="zh-cn">
<TITLE>诚信物流网 会员控制中心 - 物流商人的网上家园</TITLE>
<META qyjj=zh-cn http-equiv=qyjj-Language> 
<META qyjj="text/html; charset=gb2312" http-equiv=qyjj-Type> 
<META qyjj="Microsoft FrontPage 5.0" name=GENERATOR content="Microsoft FrontPage 5.0">
<style type="text/css">
<!--
body {
	background-color: #2C68B1;
	margin: 0px;
}
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
<link href="images/main.css" rel="stylesheet" type="text/css">
<script language="javascript">
function info(mylink)
{
window.open(mylink,'','top=50,left=120,width=495,height=310,scrollbars=yes')
}
</script> <script Language="JavaScript"> 
function FormCheck() 
{ 
var check=true;
for(var i=0;i<document.Form1.url.length;i++){
if (document.Form1.url[i].checked) {
check=false;
break;
} 
}
if (check){
alert("请选择网站模版样式！");
document.Form1.url[1].focus();
return false;
}
return true; 
document.Form1.submit()
} 
</script>
</HEAD>

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
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 商铺风格设置</td>
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
                        <div align="center"><font color="#FFFFFF">网站模版选择</font></div>
                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table>
<%if session("hydj")=1 then
    response.write("对不起您没有权限！")
else
%>            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
                <td valign="top" bgcolor="#FFFFFF"> 
                  <FORM method="POST" action="mod_save.asp" name="Form1" onSubmit="return FormCheck();" >
  <%
dim rs
dim sql
set rs=server.createobject("adodb.recordset") 
sql="select * from mod order by id desc"
rs.open sql,conn,1,1 
if rs.eof and rs.bof then 
response.write "<p align='center'>对不起，没有任何信息！</a></p>" 
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
showContent totalput,MaxPerPage
showpage totalput,MaxPerPage,"webmod.asp"
else 
if (currentPage-1)*MaxPerPage<totalPut then 
rs.move (currentPage-1)*MaxPerPage 
dim bookmark 
bookmark=rs.bookmark 
showContent totalput,MaxPerPage
showpage totalput,MaxPerPage,"webmod.asp"
else 
currentPage=1 
showContent totalput,MaxPerPage
showpage totalput,MaxPerPage,"webmod.asp"
end if 
end if 
rs.close 
end if 
set rs=nothing
%>
                    <table width="24%">
                      <tr> <TD height="1" style="font-family: 宋体; font-size: 9pt; line-height: 13pt; letter-spacing: 0pt" colspan="3" width="100%"> 
<INPUT type=hidden value=name=which> <FONT face=宋体> <p align="center"> <br> <input type="button" value="返回" onclick=javascript:history.go(-1)> 
<input type="submit" value="下一步&gt;&gt;"></FONT></TD></tr> </table></form>
<%
sub showContent (totalput,MaxPerPage)
%>
 
                  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber10">
                    <tr> 
                      <%do while not rs.eof%>
                      <TD height="1" style="font-family: 宋体; font-size: 9pt; line-height: 13pt; letter-spacing: 0pt" width="183"> 
                        <TABLE border=0 borderColorDark=#660000 
borderColorLight=#660000 cellPadding=2 cellSpacing=1 
height=90 width=160 style="border-collapse: collapse; font-family:宋体; font-size:9pt" bgcolor="#00CCFF">
                          <TBODY> 
                          <TR> 
                            <TD align=middle height=77 
onmouseout="this.style.backgroundColor =''" 
onmouseover="this.style.backgroundColor ='#0033FF'" style="font-family: 宋体; font-size: 9pt; line-height: 13pt; letter-spacing: 0pt" width="155" bgcolor="#FFFFFF"> 
                              <FONT face=宋体> <a href="javascript:info(&quot;<%=rs("src")%>&quot;)"> 
                              <IMG border=0 height=110 
src="<%=rs("src")%>" 
style="CURSOR: hand" width=150></a></FONT></TD>
                          </TR>
                          <TR bgColor=#c33313> 
                            <TD align=middle height=22 style="font-family: 宋体; font-size: 9pt; line-height: 13pt; letter-spacing: 0pt"> 
                              <FONT face=宋体> 
                              <INPUT class=f11 name=url type=radio value="<%=rs("typename")%>/index.asp?gsid=<%=gsid%>">
                              <a href="javascript:info(&quot;<%=rs("src")%>&quot;)" style="text-decoration: none"><font color="#F2F2F2"><%=rs("type1")%></font></a> 
                              </FONT> </TD>
                          </TR>
                          </TBODY>
                        </TABLE>
                        
                      </TD>
                      <% x=x+1
if x>=MaxPerPage then exit do
  if x mod 3= 0 then
  response.write "</tr><tr>"
  end if
rs.movenext 
loop 
%>
                    </tr>
                  </table>
                  <%
END SUB
Function showpage(totalnumber,maxperpage,filename) 
dim n 
if totalnumber mod maxperpage=0 then 
n= totalnumber \ maxperpage 
else 
n= totalnumber \ maxperpage+1 
end if 
response.write "<center>" 
response.write "<table border=0 width=97% cellspacing=0 cellpadding=0>" 
response.write "<tr height=10>"
response.write "<td align=left>"
response.write "共<font color=#ff6600><b>"&n&"</b></font>页&nbsp;第<font color=#ff6600><b>"&CurrentPage&"</b></font>页&nbsp;共有<font color=#ff6600><b>"&totalnumber&"</b></font>个网站模板</td>" 
response.write "<td align=right>"
if CurrentPage<2 then 
response.write "【最前页】【上一页】" 
else
response.write "【<a href="&filename&"?page=1>最前页</a>】" 
response.write "【<a href="&filename&"?page="&CurrentPage-1&">上一页</a>】&nbsp;" 
end if 
if n-currentpage<1 then 
response.write "【下一页】【最后页】" 
else 
response.write "【<a href="&filename&"?page="&(CurrentPage+1)&">" 
response.write "下一页</a>】【<a href="&filename&"?page="&n&">最后页</a>】" 
end if 
response.write "</td>"
response.write "</tr>"
response.write "</table>"
end function
%>
                </td>
            </tr>
          </table>
<%end if%>
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
</HTML>