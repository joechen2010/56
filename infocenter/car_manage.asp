<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../Member/check.asp"-->
<html>
<head>
<title><%=WebName%>-�����Ʒ��Ϣ</title>
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
<link href="../member/images/main.css" rel="stylesheet" type="text/css">
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
      <table width="100%" height="98"  border="0" cellpadding="0" cellspacing="0" align="center">
        <tr>
          <td colspan="2" align="center" valign="top">
            <table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
                <td width="100%" valign="top" bgcolor="#FFFFFF"> 
                  <table width="100%" border="0" cellpadding="4" cellspacing="1" bgcolor="#FF6500">
                    <tr bgcolor="#FFDBBF"> 
                      <td width="40" bgcolor="#FFDBBF" align="center">���</td>
                      <td width="40" bgcolor="#FFDBBF" align="center">������</td>
                      <td bgcolor="#FFCEAA" align="center" width="62">�����</td>
                      <td width="70" bgcolor="#FFCEAA" align="center">����</td>
                      <td width="53" bgcolor="#FFCEAA" align="center">��λ</td>
                      <td width="57" bgcolor="#FFCEAA" align="center">����ʱ��</td>
                      <td width="36" bgcolor="#FFCEAA" align="center">�鿴</td>
                      <td width="97" bgcolor="#FFCEAA" align="center">����</td>
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
    sql="select * from info2 where gsid='"+session("user")+"' order  by AddTime desc"
	sfilename="Pro_Manage.asp"
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> �� û �� �� �� �� Ϣ</td></tr> </p>"
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
                      <td align="center" bgcolor="#FFDBBF">
	<%if rs("infotype")="��Դ��Ϣ" then %>
      <div align="center"><img src="../images/car.gif" width="25" height="20">
      </div>
	  <%else%>
	        <div align="center"><img src="../images/goods.gif" width="25" height="20">
      </div>
	  <%end if%>					  
					  </td>
                      <td align="center" bgcolor="#FFDBBF"><%=rs("city")%></td>
                      <td bgcolor="#FFCEAA" align="center"><%=rs("city2")%></td>
                      <td bgcolor="#FFCEAA" align="center"><%=rs("cartype")%></td>
                      <td bgcolor="#FFCEAA" align="center"><%=rs("carload")%></td>
                      <td bgcolor="#FFCEAA" align="center"><%=rs("addtime")%></td>
                      <td bgcolor="#FFCEAA" align="center"><%=rs("Hits")%></td>
                      <td bgcolor="#FFCEAA" align="center"><a href="car_Edit.asp?infoID=<%=rs("infoID")%>">�޸�</a>&nbsp;<a href="car_Del.asp?infoid=<%=rs("infoID")%>&Action=Del" Onclick="return ConfirmDel();" >ɾ��</a></td>
                    </tr>
                    <% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
%>
<%end sub%>
                  </table>				  </td>
            </tr>

  <tr>
    <td height="50" colspan="2" align="center" valign="top" bgcolor="#FFFFFF" id="main">
	<a href="car_add.asp">������Դ��Ϣ</a>
	<a href="goods_add.asp">������Դ��Ϣ</a>
	</td>
    </tr>  		
</table>
</td>
</tr>
</table>

</body>
</html>

