<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }

</script>
</head>

<body>
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#DEDFDE" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="500">
  <tr align="Center" style="color:White;background-color:#6B696B;font-weight:bold;"> 		
      <td >信息类型</td>
      <td >库房类型</td>
      <td>仓库地点</td>
      <td>面积</td>
      <td>价格</td>
	  <td>发布日期</td>
      <td align="center">详情</td>
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
				
    sql="select  * from kf_info order by infoid desc"
	sfilename="right.asp"
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
      <td><%=rs("infotype")%></td>
      <td  align="center"><%=rs("kftype")%></td>
      <td align="center"><%=rs("city")%>&nbsp;<%=rs("area")%></td>
      <td align="center"><%=rs("mji")%>平方米</td>
      <td  align="center"><%=rs("prices")%>元/平方米</td>
	  <td align="center"><%=rs("addtime")%></td>
      <td><A href="javascript:openwindow('detail.asp?InfoID=<%=rs("InfoID")%>',500,400)">详情</A></td>
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
dim sql2,city,infotype,kftype
dim rs2

        if request("city")<>"地级市" then
	    city=" and city='"&request("city")&"'"
	    end if
	    if request("infotype")<>"" then
	    infotype=" and infotype='"&request("infotype")&"'"
	    end if
	    if request("area")<>"" then
	    area=" and area='"&request("area")&"'"
	    end if
	    if request("kftype")<>"" then
	    kftype=" and kftype='"&request("kftype")&"'"
	    end if		
	'city=trim(request.form("city"))
    'city2=trim(request.form("city2"))
    'area=trim(request("area"))
    'area2=trim(request("area2"))			
    sql2="select  * from kf_info where 1=1"&city&infotype&kftype&" order by infoid desc"
	'sql="select  * from ProdMain  where 1=1"&sql1&sql2&sql3&sql4&" and larcode='车辆信息' order by adddate desc"
    Set rs2= Server.CreateObject("ADODB.Recordset")
	rs2.open sql2,conn,1,1
  	if rs2.eof and rs2.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 信 息</td></tr>"
   	else
%>  

				<%do while not rs2.eof%>
	<tr align="Center" style="color:#003399;background-color:White;">		
      <td width="128"><%=rs2("infotype")%></td>
      <td  width="80" align="center"><%=rs2("kftype")%></td>
      <td width="160" align="center"><%=rs2("city")%>&nbsp;<%=rs2("area")%></td>
      <td width="168" align="center"><%=rs2("mji")%>平方米</td>
      <td width="150" align="center"><%=rs2("prices")%>元/平方米</td>
	   <td width="150" align="center"><%=rs2("addtime")%></td>
      <td width="30"><A href="javascript:openwindow('detail.asp?InfoID=<%=rs2("InfoID")%>',500,400)">详情</A></td>
	</tr>
<% 

rs2.movenext 
loop 

%>
<%end if%>

<%end if%>	
</table>
</body>
</html>
