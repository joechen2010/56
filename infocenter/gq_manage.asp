<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../Member/check.asp"-->
<html>
<head>
<title><%=WebName%>-管理产品信息</title>
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
<link href="../Member/images/main.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
function ConfirmDel()
{
   if(confirm("确定要删除该项记录吗？        \n点击确定删除！\n点击取消返回！"))
     return true;
   else
     return false;
	 
}
</script>
</head> 
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"> 
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="566" align="center" valign="top" bgcolor="#FFFFFF" id="main">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><a href="http://www.cx56w.com">诚信物流网</a> &gt; 已发布的供求信息</td>
          <td width="17%">&nbsp;</td>
      </tr>
    </table>
	  <table width="100%" height="127"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top"><table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
                <td valign="top" bgcolor="#FFFFFF"> 
                  <table width="520" border="0" cellpadding="4" cellspacing="1" bgcolor="#FF6500">
                    <tr bgcolor="#FFDBBF"> 
                      <td width="81" bgcolor="#FFDBBF" align="center">信息类型</td>
                      <td bgcolor="#FFCEAA" align="center" width="111">信息标题</td>
                      <td width="83" align="center" bgcolor="#FFCEAA">仓库地点</td>
                      <td width="57" bgcolor="#FFCEAA" align="center">发布时间</td>
                      <td width="36" bgcolor="#FFCEAA" align="center">查看</td>
                      <td width="97" bgcolor="#FFCEAA" align="center">操作</td>
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
strUnit="个信息"
    sql="select * from gq_info where gsid='"+session("user")+"' order  by AddTime desc"
	sfilename="gq_Manage.asp"
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 信 息</td></tr> </p>"
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
                      <td bgcolor="#FFCEAA" align="center"><%=rs("title")%></td>
                      <td align="center" bgcolor="#FFCEAA"><%=rs("city")%></td>
                      <td bgcolor="#FFCEAA" align="center"><%=rs("addtime")%></td>
                      <td bgcolor="#FFCEAA" align="center"><%=rs("Hits")%></td>
                      <td bgcolor="#FFCEAA" align="center"><a href="gq_Edit.asp?infoID=<%=rs("infoID")%>">修改</a>&nbsp;<a href="gq_Del.asp?infoid=<%=rs("infoID")%>&Action=Del" Onclick="return ConfirmDel();" >删除</a></td>
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
                  </table>                </td>
            </tr>
          </table>
         <table width="534"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="center"><a href="gq_add.asp">发布供求</a></td>
              </tr>
            </table>
		</td>
       </tr>
     </table>
    </td>
  </tr>
</table>

</body></html>
