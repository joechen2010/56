<%@ codepage ="936" %>
<%
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache"
%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<%
const MaxPerPage=4
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
dim rstype
mysql="select * from qyml where user='"+session("user")+"'" 
Set myrs= Server.CreateObject("ADODB.Recordset") 
myrs.open mysql,conn,1,1
if myrs.eof and myrs.bof then
    response.write "<script>alert('�Բ��𣬸��û�������.');top.location.href='../login/login.asp';</script>"
	response.end
end if 
%> 
<meta http-equiv="Content-Language" content="zh-cn">
<TITLE>���������� ��Ա�������� - �������˵����ϼ�԰</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
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
</HEAD>
<BODY text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
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
            <a href="http://www.cx56w.com">����������</a> &gt; ���յ�������</td>
          <td width="17%"><img src="images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="534"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> <%=session("user")%> </font> </strong>����ӭ��������</td>
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
                        <div align="center"><font color="#FFFFFF">���յ�������</font></div>
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
                  <% 
sql="select * from Message where T_User='"&session("user")&"' order by id desc"
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
if rs.eof and rs.bof then 
   response.write "������Ϣ!"
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
else 
if (currentPage-1)*MaxPerPage<totalPut then 
rs.move (currentPage-1)*MaxPerPage 
dim bookmark 
bookmark=rs.bookmark 
showContent
else 
currentPage=1 
showContent
end if 
end if
end if 
sub showContent 
%>
              <table border=0 cellpadding=0 cellspacing=0 width="100%" style="border-collapse: collapse" bordercolor="#111111" height="122">
                <tr> 
                      <td valign=top height="122" width="100%"> 
                        <table border=0 cellpadding=5 cellspacing=1 width="100%" style="border-collapse: collapse" bgcolor="#B1D4F2">
                          <tr align="center"> 
                            <td width="25%" height="20" background="images/title_bg_2.gif">����</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">������</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">����ʱ��</td>
                            <td width="25%" height="20" background="images/title_bg_2.gif">����</td>
                          </tr>
                          <%do while not rs.eof%>
                          <tr> 
                            <td height="19" bgcolor="#EDF6FF"><a href="read.asp?id=<%=rs("id")%>"><%=rs("subject")%></a> 
                              <%if rs("new")=0 then response.write "(<font color=#ff000>New</font>)"%>
                            </td>
                            <td height="19" bgcolor="#EDF6FF"><%=rs("F_Name")%></td>
                            <td height="19" bgcolor="#EDF6FF"><%=rs("TF_Time")%></td>
                            <td height="19" bgcolor="#EDF6FF"><a href="reply.asp?id=<%=rs("id")%>">�ظ�</a> 
                              <a href="msg_del.asp?action=del&id=<%=rs("id")%>&page=<%=currentpage%>">ɾ��</a></td>
                          </tr>
                          <% i=i+1 
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
rs.close 
set rs=nothing 
%>
                        </table>
                    <table border=0 cellpadding=5 cellspacing=1 width="100%">
                      <tr> 
                        <td colspan="2"> 
                          <%showpage totalput,MaxPerPage,"inbox.asp"%>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
              <%
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
response.write "<font color=red><b>��"&n&"ҳ&nbsp;��"&CurrentPage&"ҳ " 
response.write "����"&totalnumber&"��������Ϣ</b></font>" 
response.write "</td>"
response.write "<td VALIGN=middle align=right >"
response.write "����ǰҳ������һҳ��" 
else 
response.write "<font color=red><b>��"&n&"ҳ&nbsp;��"&CurrentPage&"ҳ " 
response.write "����"&totalnumber&"��������Ϣ</b></font>" 
response.write "��<a href="&filename&"?page=1&id="&id&">��ǰҳ</a>��" 
response.write "��<a href="&filename&"?page="&CurrentPage-1&"&id="&id&">��һҳ</a>��" 
end if 
if n-currentpage<1 then 
response.write "����һҳ�������ҳ��" 
else 
response.write "��<a href="&filename&"?page="&(CurrentPage+1)&"&id="&id&">" 
response.write "��һҳ</a>���� <a href="&filename&"?page="&n&"&id="&id&">���ҳ</a>��" 
end if 
response.write "</form>" 
end function
%>
</td>
            </tr>
          </table>
                        <table width="534"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="right"><a href="#top"><img src="images/icon_top.gif" alt="�ص�ҳ�涥��" border="0" align="absmiddle" width="15" height="18"></a> 
                  <a href="#top">�ټ��һ��</a></td>
              </tr>
            </table></td>
        </tr>
      </table>
</td>
  </tr>
</table>

<!--#include file="bottom.htm"-->
</BODY>
</HTML>
<%myrs.close 
set myrs=nothing
conn.close
set conn=nothing
%>