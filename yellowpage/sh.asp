<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<html>
<head>
<title>������������������ҳ</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css 
rel=stylesheet>
</head>
<body>
<!--#include file="../inc/top.htm"-->
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="center"><iframe id="ifrsearch" src="../search/index_login.asp" width="100%" frameborder="0" scrolling="no" height="90"></iframe></td>
  </tr>
</table>
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<table width="778" border="0" cellpadding="0" cellspacing="1" bgcolor="#FF6600" align="center">
  <tr bgcolor="FFC390"> 
    <td height="2" colspan="4"></td>
  </tr>
  <tr bgcolor="#FF6600"> 
    <td width="66%" height="26" class="text"><font color="#FFFFFF">������ҳ</font></td>
    <td width="34%" height="24" align="right" class="text">&nbsp;</td>
  </tr>
</table>
<table align="center" width="778" cellpadding="0" cellspacing="0" border="0">
  <tr> 
    <td align="center" width="52%" valign="top"> 
      <table width="99%" border="0" cellspacing="1" cellpadding="0" bgcolor="#FF9900" align="left">
        <tr>
          <td bgcolor="#FFFFFF" align="center"><a href="index.asp">���ص�ȫ����ͼ</a></td>
        </tr>	  
        <tr>
          <td bgcolor="#FFFFFF" height="396"><img src="images/102.gif" width="520" height="360"  border="0" usemap="#Map"></td>
        </tr>
      </table>
      
    </td>
    <td align="center" width="48%" valign="top">
	<table cellspacing="0" cellpadding="2" align="Center" bordercolor="Tan" border="0" id="DataGrid1" style="color:Black;background-color:LightGoldenrodYellow;border-color:Tan;border-width:1px;border-style:solid;font-size:X-Small;width:100%;border-collapse:collapse;">
	<tr align="Center" style="background-color:Tan;font-weight:bold;">
		<td align="center">����</td><td align="center">������</td><td align="center">��λ����</td>
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
	if request("province")<> "" then			
    sql="select * from qyml where province='"&request("province")&"' order by id desc"
	else
	 if request("city")<>"" then
    sql="select * from qyml where city='"&request("city")&"' order by id desc"
	  else
	sql="select * from qyml order by id desc"
	end if
	end if
	sfilename="index.asp"
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> �� û �� �� �� �� Ա</td></tr> </p>"
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

	
sub showContent 
dim i
i=0
%>  

				<%do while not rs.eof%>	
	<tr align="Center" onMouseOver="this.style.backgroundColor='#efefef'" onMouseOut="this.style.backgroundColor=''">
		<td><%=rs("qylb")%></td>
		<td><%=rs("city")%></td>
		<td><a href="../yellowpage/site.asp?InfoID=<%=rs("ID")%>&user=<%=rs("user")%>" target="_blank"><%=rs("qymc")%></a></td>
	</tr>
<% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 

%>
<%end sub%> 


</table>
</td>
  </tr>
</table>
<!--#include file="../inc/down.htm"-->
<map name=Map> 
  <area shape=rect alt=¬���� coords=258,182,292,199 href=sh.asp?city=%C2%AC%CD%E5%C7%F8 title="¬����" >
	
  <area shape=rect alt=������ coords=165,165,206,180 href=sh.asp?city=%B3%A4%C4%FE%C7%F8 title="������" >
	
  <area shape=rect alt=�ֶ��� coords=361,208,399,224 href=sh.asp?city=%C6%D6%B6%AB%C7%F8 title="�ֶ���" >
	
  <area shape=rect alt=������ coords=116,257,159,272 href=sh.asp?city=%E3%C9%D0%D0%C7%F8 title="������" >
  <area shape="rect" coords="64,313,104,328" href="sh.asp?city=%BD%F0%C9%BD%C7%F8" alt="��ɽ��" title="��ɽ��">
  <area shape="rect" coords="158,340,198,354" href="sh.asp?city=%B7%EE%CF%CD%C7%F8" alt="������" title="������">
  <area shape="rect" coords="410,325,455,339" href="sh.asp?city=%C4%CF%BB%E3%C7%F8" alt="�ϻ���" title="�ϻ���">
	
  <area shape=rect alt=������ coords=273,159,309,174 href=sh.asp?city=%BB%C6%C6%D6%C7%F8 title="������" >
	
  <area shape=rect alt=�ӱ� coords=6,134,46,151 href=sh.asp?city= >
	
  <area shape=rect alt=�ζ��� coords=91,101,133,120 href=sh.asp?city=%BC%CE%B6%A8%C7%F8 title="�ζ���" >
	
  <area shape=rect alt=����� coords=285,83,328,101 href=sh.asp?city=%BA%E7%BF%DA%C7%F8 title="�����" >
	
  <area shape=rect alt=�ɽ��� coords=18,243,60,259 href=sh.asp?city=%CB%C9%BD%AD%C7%F8 title="�ɽ���" >
	
  <area shape=rect alt=����� coords=193,225,233,243 href=sh.asp?city=%D0%EC%BB%E3%C7%F8 title="�����" >
	
  <area shape=rect alt=������ coords=234,146,267,162 href=sh.asp?city=%BE%B2%B0%B2%C7%F8 title="������" >
	
  <area shape=rect alt=������ coords=348,91,391,110 href=sh.asp?city=%D1%EE%C6%D6%C7%F8 title="������" >
	
  <area shape=rect alt=������ coords=197,111,238,128 href=sh.asp?city=%C6%D5%CD%D3%C7%F8 title="������" >
	
  <area shape=rect alt=բ���� coords=238,58,283,76 href=sh.asp?city=%D5%A2%B1%B1%C7%F8 title="բ����" >
	
  <area shape=rect alt=��ɽ�� coords=179,26,225,44 href=sh.asp?city=%B1%A6%C9%BD%C7%F8 title="��ɽ��" >
	
  <area shape=rect alt=������ coords=274,5,312,23 href=sh.asp?city=%B3%E7%C3%F7%CF%D8 title="������" >
</map>

</body>
</html>

