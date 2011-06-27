<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
province=request("province")
infotype=request("infotype")
city=request("city")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }

</script>
</head>
<body>
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#3366CC" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="500">
    <tr align="Center" style="color:#CCCCFF;background-color:#003399;font-weight:bold;">		
      <td >标题</td>
      <td >所在省市</td>
      <td >发表日期</td>
      <td  align="center">详情</td>
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
        if request("city")<>"地级市" then
	    c=" and province='"&request("province")&"' and city='"&request("city")&"'"
		 else
		 if request("province")<>"省份" then
		 c=" and province='"&request("province")&"'"
		 end if
	    end if
	    if request("area")<>"" then
	    area=" and area='"&request("area")&"'"
	    end if
	    if request("infotype")<>"" then
	    itype=" and infotype='"&request("infotype")&"'"
	    end if				
    sql="select  * from gq_info where 1=1"&c&itype&" order by infoid desc"
	sfilename="search.asp?province="&request("province")&"&city="&request("city")&"&infotype="&request("infotype")&""
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
      <td width="130">
				  <%if len(rs("title"))>8 then%>
				  <%=left(rs("title"),6)%>...
				  <%else%>
				 <%=rs("title")%>
				  <%end if%>	  
	  </td>
      <td  width="299" align="center">&nbsp;<%=rs("city")%>&nbsp;<%=rs("area")%></td>
      <td width="182" align="center"><%=rs("addtime")%></td>
      <td width="125"><A href="javascript:openwindow('detail.asp?InfoID=<%=rs("InfoID")%>',500,400)">详情</A></td>
	</tr>
<% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
%>
<%end sub %> 
</table>
</body>
</html>
