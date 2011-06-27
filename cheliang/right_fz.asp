
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<HTML>
<HEAD>
<title>专线信息检索条件</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css 
rel=stylesheet>
<STYLE type=text/css>.bg {
	BACKGROUND-POSITION: 50% top; BACKGROUND-REPEAT: no-repeat
}
</STYLE>
<script language="javascript">
function openwindow(url,width,height) 
{  left1 = (screen.width-800)/2; 	
   top1 = (screen.height-350)/2; 	
   window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); 
}

</script>
</HEAD>

<body>
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#DEDFDE" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="500">
  <tr align="Center" style="color:White;background-color:#6B696B;font-weight:bold;"> 		
      <td style="width:130px;" width="130">车牌号</td>
      <td width="149">车籍</td>
      <td width="150">车主</td>
      <td style="width:160px;" width="182">车型</td>
      <td style="width:30px;" width="125" align="center">详情</td>
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
        if session("fz")<>"" then
        fz=" and (b.province='"&session("fz")&"' or b.city='"&session("fz")&"')"
		end if				
    sql="select * from file_info a,qyml b where 1=1"&fz&" and a.gsid=b.user order by infoid desc"
	sfilename="right_fz.asp"
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
      <td width="130"><%=rs("carnum")%></td>
      <td  width="149" align="center"><%=rs("city")%></td>
      <td  width="150" align="center">
				  <%if len(rs("qymc"))>8 then%>
				  <%=left(rs("qymc"),6)%>...
				  <%else%>
				 <%=rs("qymc")%>
				  <%end if%>	  
	  </td>
      <td width="182" align="center"><%=rs("cartype")%></td>
      <td width="125"><A href="javascript:openwindow('detail.asp?InfoID=<%=rs("InfoID")%>',500,400)">详情</A></td>
	</tr>
                    <% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
%>
<%end sub%> 

 

</table>
</body>
</html>
