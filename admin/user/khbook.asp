<%@ codepage ="936" %>
<%
if instr(session("flag"),"01")=6 then
response.redirect "../login.asp"
response.end
end if
%>
<!--#include file="../../Inc/Conn.asp" -->
<HTML><HEAD>
<TITLE>在线咨询</TITLE> <META content=zh-cn http-equiv=Content-Language> <META content="text/html; charset=gb2312" http-equiv=Content-Type> 
<LINK href="../../images/style.css" rel=stylesheet type=text/css> <STYLE type=text/css>
<!--
td { font-family: "宋体"; font-size: 9pt; line-height: 14pt; }
-->
</STYLE> <META content="Microsoft FrontPage 5.0" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginwidth="0" marginheight="0"> 
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%" style="border-collapse: collapse" bordercolor="#111111"> 
<TBODY> <TR> <TD vAlign=top width="100%"> <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%" style="border-collapse: collapse" bordercolor="#111111"> 
<TBODY> <TR> <TD align=right width=38> 　</TD><TD height=25 width=52><B><FONT 
color=#cc6600>　</FONT></B><IMG 
src="../../Guestbook/images/x_4.gif"> </TD><TD height=25 width="100%"> <span style="font-size: 10.5pt; letter-spacing: 1"><b>在线咨询管理</b></span></TD></TR></TBODY></TABLE><div align="center"> 
<center> <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%" style="border-collapse: collapse" bordercolor="#111111"> 
<TBODY> <TR> <TD bgColor=#01438b height=1></TD></TR> <TR> <TD bgColor=#217dda height=5></TD></TR></TBODY></TABLE></center></div><center> 
<TABLE border=0 cellSpacing=1 width="100%" style="border-collapse: collapse" bordercolor="#111111" background="../../Guestbook/images/bg1.gif"> 
<TBODY> <TR> <TD align=middle height="12" width="100%">　</TD></TR> <TR> <TD align=middle height="169" width="100%"> 
<TABLE border=0 cellPadding=0 cellSpacing=0 height=112 width="100%" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#FFFFFF"> 
<TBODY> <TR vAlign=top> <TD colSpan=5 height="112"> <div align="center"> <center> 
<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111"> 
<tr> <td valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" height="132"> 
<tr> <td height="112"> <%
const MaxPerPage=8 '分页显示留言个数
dim sql,rs,totalput,CurrentPage,TotalPages,i
if not isempty(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if
%> <%
set rs=server.createobject("adodb.recordset")
sql="select * from GuestBook order by id desc"
rs.open sql,conn,1,1 
if rs.eof and rs.bof then 
response.write "<p align='center'> 还 没 有 任 何 留 言</p>" 
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
showpages 
else 
if (currentPage-1)*MaxPerPage<totalPut then 
rs.move (currentPage-1)*MaxPerPage 
dim bookmark 
bookmark=rs.bookmark 
showContent 
showpages 
else 
currentPage=1 
showContent 
end if 
end if 
rs.close 
end if 
set rs=nothing 
conn.close 
set conn=nothing 
sub showContent 
i=0 
do while not (rs.eof or err) %> <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="border-collapse: collapse"> 
<tr > <td width="95%" > <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111"> 
<tr> <td width="485" align="right" valign="bottom"> <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber10" background="../../Guestbook/images/bg23.gif" height="20"> 
<tr> <td width="100%"> <b><font color="#000066"> &nbsp; <img src="../../Guestbook/images/gb.gif"> 
主题：&nbsp;</font><font color="#000099"><%=rs("Title")%></font></b><font color="#CC6600">&nbsp;&nbsp;&nbsp; 
</font> <%if rs("where")<>"" then%> <IMG alt="<%=rs("where")%>" border=0 
src="../../Guestbook/images/gb_where.gif"> <%else%> <IMG alt="未知" border=0 
src="../../Guestbook/images/l1.gif"> <%end if%> <%if rs("tel")<>"" then%> <IMG alt="<%=rs("tel")%>" border=0 
src="../../Guestbook/images/gb_msn.gif"> <%else%> <IMG alt="未知" border=0 
src="../../Guestbook/images/l2.gif"> <%end if%> <%if rs("UserEmail")="@" then%> 
<IMG 
alt="未知" border=0 src="../../Guestbook/images/l3.gif"> <%else%> <a href="mailto:<%=rs("UserEmail")%>"> 
<IMG 
alt="<%=rs("UserEmail")%>" border=0 src="../../Guestbook/images/gb_mail.gif"></a> 
<%end if%> </td></tr> </table></td><td valign="bottom" width="110">　</td><td width="36" align="center" valign="bottom"> 
<a href="hfbook.asp?id=<%=rs("Id")%>">回复</a></td><td width="14" align="center" valign="bottom">　</td><td width="43" align="center" valign="bottom"> 
<a href="Del.asp?id=<%=rs("Id")%>">删除</a>&nbsp;</td></tr> </table></td></tr> <tr> 
<td> <div align="center"> <center> 
                                              <table width="100%" border="0" cellpadding="6" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" height="62">
                                                <tr> <td width="89" align="right" valign="top" height="50" rowspan="2" bgcolor="#FDFEFF"> 
<strong><font color="#000066"><br> 内容：&nbsp; </font></strong></td><td height="29" width="648" valign="middle" colspan="2" bgcolor="#FDFEFF"> 
<div align="center"> <center> <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="98%" id="AutoNumber9"> 
<tr> <td width="100%"> <%=rs("Content")%></td></tr> </table></center></div></td><td height="50" width="41" rowspan="2" bgcolor="#FDFEFF"> 
　</td></tr> <tr> <td height="21" width="364" bgcolor="#FDFEFF"> 　</td><td height="21" width="284" bgcolor="#FDFEFF"> 
<p><font color="#CC6600"><b>&nbsp; →</b><a href="mailto:<%=rs("UserEmail")%>"><%=rs("UserName")%></a></font><font color="#808080">[<%=rs("AddTime")%>]</font></td></tr> 
<%
set rsg=server.createobject("adodb.recordset")
sql="select * from hfbook where Gid="&rs("Id")
rsg.open sql,conn,1,1 
if rsg.eof then
response.Write(" ") 
else
while not rsg.eof
%> <tr> <TD colSpan=4 height="12" width="778"> <div align="center"> <center> <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="95%" id="AutoNumber7"> 
<tr> <TD height=1 background="../../Guestbook/images/line.gif"></TD></tr> </table></center></div></TD></tr> 
</table></center></div></td></tr> <tr> <td> <table width="100%" border="0" cellpadding="0" cellspacing="0"> 
<tr> <td width="70" align="right"><strong> <font color="#FF3300"> 回复：&nbsp;</font></strong></td><td height="20" bgcolor="#FDFCF2"> 
<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111"> 
<tr> <td width="665" height="20" bgcolor="#FDFBEE"> <%If rsg("IsA")=True then%> 
<font color="#FF3300">&lt;&gt; <%=rsg("Gcontent")%> —— </font> <%if rsg("GEmail")="@" then%> 
<font color="#FF3300"><b><%=rsg("Gname")%></b></font> <%else%> <font color="#FF3300"> 
<b><%=rsg("Gname")%></b> </font> <a href="mailto:<%=rsg("GEmail")%>"><font color="#FF3300"> 
[Email]</font></a> <%end if%> <font color="#FF3300"><%=rsg("Gtime")%></font> <%else%> 
<font color="#000066">&lt;&gt; <%=rsg("Gcontent")%> —— </font> <%if rsg("GEmail")="@" then%> 
<font color="#000066"><b><%=rsg("Gname")%></b></font> <%else%> <font color="#000066"><b><%=rsg("Gname")%></b></font> 
<a href="mailto:<%=rsg("GEmail")%>"><font color="#000066"> [Email]</font></a> 
<%end if%> <font color="#000066"><%=rsg("Gtime")%></font> <%End If%><font color="#FF0000">[<a href="Del.asp?gid=<%=rsg("Id")%>" style="text-decoration: none"><font color="#FF3300">删除</font></a>]</font></td><td width="43" align="center"> 
　</td></tr> </table><%
rsg.movenext
wend
rsg.close
set rsg=nothing
end if
%> </td></tr> </table></td></tr> </table></td></tr> <tr> <td align="right" bgcolor="#EEF3FF" height="1"> 
<% i=i+1 
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
end sub 
sub showpages() 
dim n 
if (totalPut mod MaxPerPage)=0 then 
n= totalPut \ MaxPerPage 
else 
n= totalPut \ MaxPerPage + 1 
end if 
if n=0 then 
response.write "<p align='right'><a href=gbadd.asp>请留下您宝贵的意见！</a></p>" 
exit sub 
end if 
dim k 
response.write "分页 &gt;&gt; " 
for k=1 to n 
if k=currentPage then 
response.write"["+Cstr(k)+"]" 
else 
response.write"<a href='khbook.asp?page="+cstr(k)+"'>[ "+Cstr(k)+" ]</a>"
end if 
next 
end sub 
%> </td></tr> <tr> <td height="20" align="right" valign="bottom" bgcolor="#217DDA"> 
　</td></tr> </table></td></tr> </table></center></div></TD></TR></TBODY></TABLE></TD></TR> 
</TBODY></TABLE></center></TD><TD width=3></TD></TR> </TBODY></TABLE> 
</BODY>
</HTML>