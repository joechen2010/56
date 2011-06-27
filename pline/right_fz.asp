<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
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
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#DEDFDE" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="500">
  <tr align="Center" style="color:White;background-color:#6B696B;font-weight:bold;"> 
    <td style="width:130px;" width="120">专线名称</td>
    <td width="203">途经(中转)城市</td>
    <td style="width:160px;" width="139">企业名称</td>
    <td style="width:30px;" width="50" align="center">专线</td>
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
        fz=" and (a.province='"&session("fz")&"' or a.city='"&session("fz")&"')"
		end if
	 sql="select  * from zx_info a,qyml b where 1=1"&fz&" and infotype='客运专线' and a.gsid=b.user order by infoid desc"
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
    <td><%=rs("a.city")%>←→<%=rs("city2")%></td>
    <td  width="203" align="center"> 
      <%
						   set rs1=server.CreateObject("adodb.recordset")
						    sql1 = "select * from zx_info2 where infoid="&rs("infoid")&" order by id desc"
						   rs1.open sql1,conn,1,1
						   if not rs1.eof then
						    for j=1 to rs1.recordcount
						   %>
      <%=rs1("city")%>&nbsp; 
      <%
							  rs1.movenext
							  next
							  end if
							  %>
    </td>
    <td width="139" align="center"> 
      <%=rs("qymc")%> 
    </td>
    <td><A href="javascript:openwindow('detail.asp?InfoID=<%=rs("InfoID")%>',500,400)">详情</A></td>
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
