<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<TITLE>���������� ��Ա�������� - �������˵����ϼ�԰</TITLE>
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
<script language="JavaScript" type="text/JavaScript">
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ�������¼��        \n���ȷ��ɾ����\n���ȡ�����أ�"))
     return true;
   else
     return false;
	 
}
</script>
</head>
<BODY text=#000000 leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" bgcolor="#F7F7FF">
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
            <a href="http://www.cx56w.com">����������</a> &gt; ��ҵ����֤��</td>
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
                        <div align="center"><font color="#FFFFFF">��ҵ����֤��</font></div>
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
sfilename="qyry.asp"
strUnit="������"
const MaxPerPage=20
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
sql="select * from qyzs where gsid='"+session("user")+"' order by ID desc"
set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
if rs.eof and rs.bof then
    response.write "<p>&nbsp;</p><p align=center>����ʱ��û�з�����Ϣ��<p>"
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
call showpage2(sfilename,totalPut,maxperpage,true,true,strUnit)
else 
if (currentPage-1)*MaxPerPage<totalPut then 
rs.move (currentPage-1)*MaxPerPage 
dim bookmark 
bookmark=rs.bookmark 
showContent 
call showpage2(sfilename,totalPut,maxperpage,true,true,strUnit)
else 
currentPage=1 
showContent 
call showpage2(sfilename,totalPut,maxperpage,true,true,strUnit)
end if 
end if 
end if
conn.close
set conn=nothing 
sub showContent
%>
                  <table border="0" cellspacing="0" width="100%" style="border-collapse: collapse" height="1">
                    <tr height=30 bgcolor="#0066CC"> 
                      <td width="201" height="23"><font color="#FFFFFF">&nbsp;֤������</font></td>
                      <td width="205" height="12"><font color="#FFFFFF">&nbsp;�ϴ�ʱ��</font></td>
                      <td width="110" height="12">&nbsp;</td>
                    </tr>
                    <%do while not rs.eof%>
                    <tr> 
                      <td width="201">&nbsp;<%=rs("zsmc")%></td>
                      <td width="205">&nbsp;<%=rs("scsj")%></td>
                      <td width="110"><a href="Qyry_del.asp?id=<%=rs("id")%>" OnClick="return ConfirmDel();">ɾ��</a>|<a href="qyry_modfiy.asp?id=<%=rs("id")%>">�޸�</a></td>
                    </tr>
                    <% i=i+1 
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
rs.close 
set rs=nothing 
%>
                  </table>
                  <%
end sub 
%>
                 <br>
                  <div align="right"><a href="qyry_add.asp">�ϴ���ҵ����֤��</a> </div>
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
</body>
</html>