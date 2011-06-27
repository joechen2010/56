
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<HTML>
<HEAD>
<title>专线信息检索条件</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<STYLE type=text/css>.bg {
	BACKGROUND-POSITION: 50% top; BACKGROUND-REPEAT: no-repeat
}
</STYLE>
</HEAD>
<body leftmargin="0" rightmargin="0" style="FONT-SIZE: 10pt" bgColor="#FFFFFF" topmargin="0">
<!--#include file="../inc/top_beijing.htm" -->
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="center"><iframe id="ifrsearch" src="../search/index_job.asp" width="100%" frameborder="0" scrolling="no" height="90"></iframe></td>
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
    <td width="66%" height="26" class="text"><font color="#FFFFFF">招聘信息</font></td>
    <td width="34%" height="24" align="right" class="text">&nbsp;</td>
  </tr>
</table>
<table width="778" border="0" align="center" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#3366CC" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="778">
    <tr align="Center" style="color:#CCCCFF;background-color:#003399;font-weight:bold;">		
      <td style="width:130px;" width="127">招聘职位</td>
      <td width="208">招聘单位</td>
      <td style="width:160px;" width="204">所在省市</td>
      <td style="width:160px;" width="157">发布日期</td>
      <td style="width:30px;" width="30" align="center">详情</td>
	</tr>
		  <%if request("action")="" then%>
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
				
    sql="select * from zp_info a,qyml b where infotype='招聘信息' and a.gsid=b.user and a.city='北京' order by infoid desc"
	sfilename="index_z.asp"
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

	
sub showContent 
dim i
i=0
%>  

				<%do while not rs.eof%>		
	
	<tr align="Center" style="color:#003399;background-color:White;">		
      <td width="127"><%=rs("job")%></td>
      <td  width="208" align="center"><%=rs("qymc")%></td>
      <td width="204" align="center"><%=rs("a.province")%>&nbsp;<%=rs("a.city")%></td>
      <td width="157" align="center"><%=rs("addtime")%></td>
      <td width="30"><A href="detail.asp?InfoID=<%=rs("InfoID")%>">详情</A></td>
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
<%end if%> 
<%if request("action")="search" then%>

                    <%
dim sql2,province,jobtype,area,area2
dim rs2

        if request("job")<>"" then
	    job=" and job='"&request("job")&"'"
	    end if	
        if request("jobtype")<>"" then
	    jobtype=" and jobtype='"&request("jobtype")&"'"
	    end if	
		 if request("infotype")<>"" then
	    infotype=" and infotype='"&request("infotype")&"'"
	    end if		   
		 if request("city")<>"" then
	    city=" and city='"&request("city")&"'"
	    end if	
	'city=trim(request.form("city"))
    'city2=trim(request.form("city2"))
    'area=trim(request("area"))
    'area2=trim(request("area2"))			
    sql2="select  * from zp_info where 1=1"&city&jobtype&infotype&city&"  order by infoid desc"
	'sql="select  * from ProdMain  where 1=1"&sql1&sql2&sql3&sql4&" and larcode='车辆信息' order by adddate desc"
    Set rs2= Server.CreateObject("ADODB.Recordset")
	rs2.open sql2,conn,1,1
  	if rs2.eof and rs2.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 信 息</td></tr>"
   	else
%>  

				<%do while not rs2.eof%>
	<tr align="Center" style="color:#003399;background-color:White;">		
      <td width="130"><%=rs2("infotype")%></td>
      <td  width="299" align="center"><%=rs2("province")%>&nbsp;<%=rs2("city")%></td>
      <td width="182" align="center"><%=rs2("contact")%></td>
      <td width="125"><A href="detail.asp?InfoID=<%=rs2("InfoID")%>">详情</A></td>
	</tr>
<% 

rs2.movenext 
loop 

%>
<%end if%>

<%end if%>	
</table>
	
	
	</td>
  </tr>
</table>
<!--#include file="../inc/down.htm"-->
</body>
</HTML>
