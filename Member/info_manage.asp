<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/config.asp"-->
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
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
<link href="images/main.css" rel="stylesheet" type="text/css">
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
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"> 
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
            <a href="http://www.cx56w.com">����������</a> &gt; ��������Ϣ</td>
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
                        <div align="center"><font color="#FFFFFF">�ѷ�����Ϣ</font></div>
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
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>
                        <table width="210" border="0" align="left" cellpadding="4" cellspacing="1" bgcolor="#FF6500">
                          <tr bgcolor="#FFDBBF"> 
                            <td align="center" bgcolor="#FFE7C1"><a href="info_manage.asp">������Ϣ</a></td>
                            <td align="center" bgcolor="#FFE7C1"><a href="info_manage.asp?infotype=1">��Ӧ��Ϣ</a></td>
                            <td align="center" bgcolor="#FFC49B"><a href="info_manage.asp?infotype=2">����Ϣ</a></td>
                          </tr>
                        </table>    
      </td>
                    </tr>
					<tr>
					<td>&nbsp;</td>
					</tr>
                  </table>
                  <table width="520" border="0" cellpadding="4" cellspacing="1" bgcolor="#FF6500">
                    <tr bgcolor="#FFDBBF"> 
                      <td width="30" bgcolor="#FFDBBF" align="center">����</td>
                      <td bgcolor="#FFCEAA" align="center" width="200">����</td>
                      <td width="70" bgcolor="#FFCEAA" align="center">����ʱ��</td>
                      <td width="70" bgcolor="#FFCEAA" align="center">����ʱ��</td>
                      <td width="40" bgcolor="#FFCEAA" align="center">�鿴</td>
                      <td width="120" bgcolor="#FFCEAA" align="center">����</td>
                    </tr>
                    <%
    const MaxPerPage=15
   	dim totalPut   
   	dim CurrentPage
   	dim TotalPages
   	dim i,j
   	dim idlist
   	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if
dim sql
dim rs
dim qymc
dim strUnit,sfilename
strUnit="����Ϣ"
if request("inotype")="1" then
    sql="select * from info where gsid='"+session("user")+"' and infotype='��Ӧ' order by AddTime desc"
    sfilename="info_manage.asp?infotype=1"
elseif request("infotype")="2" then
    sql="select * from info where gsid='"+session("user")+"' and infotype='��' order by AddTime desc"
    sfilename="info_manage.asp?infotype=2"
else
    sql="select * from info where gsid='"+session("user")+"' order by AddTime desc"
	sfilename="info_manage.asp"
end if
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""6"" bgcolor=""#FFCEAA""> �� û �� �� �� �� Ϣ</td></tr> </p>"
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
            		call showpage2(sfilename,totalPut,maxperpage,false,false,strUnit)
       		else
          		if (currentPage-1)*MaxPerPage<totalPut then
            			rs.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs.bookmark
            			showContent
             			call showpage2(sfilename,totalPut,maxperpage,false,false,strUnit)
        		else
	        		currentPage=1
           			showContent
           			call showpage2(sfilename,totalPut,maxperpage,false,false,strUnit)
	      		end if
	   	end if
   	rs.close
   	end if
	set rs=nothing  
   	conn.close
   	set conn=nothing
	
sub showContent 
dim i
i=0
%>
                    <%do while not rs.eof%>
                    <tr bgcolor="#FFDBBF"> 
                      <td bgcolor="#FFDBBF" align="center"><%=rs("infotype")%></td>
                      <td bgcolor="#FFCEAA"><%=left(rs("title"),18)%></td>
                      <td bgcolor="#FFCEAA" align="center"><%=rs("addtime")%></td>
                      <td bgcolor="#FFCEAA" align="center"><%=rs("addtime")+rs("validate")%></td>
                      <td bgcolor="#FFCEAA" align="center"><%=rs("Hits")%></td>
                      <td bgcolor="#FFCEAA" align="center"><a href="Info_Edit.asp?InfoID=<%=rs("InfoID")%>">�޸�</a> <a href="Info_Del.asp?InfoID=<%=rs("InfoID")%>&action=Update">����</a> <a href="Info_Del.asp?InfoID=<%=rs("InfoID")%>&Action=Del" Onclick="return ConfirmDel();" >ɾ��</a></td>
                    </tr>
                    <% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
%>
                    <%
end sub 
%>
                  </table>
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

<!--#include file="bottom.htm"--></body></html>