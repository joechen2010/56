<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<HTML>
<HEAD>
<title>专线信息检索条件</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/css.css" type=text/css rel=stylesheet><!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.STYLE1 {
	color: #808080;
	font-size: 12px;
}
.STYLE2 {
	font-size: 14px;
	font-weight: bold;
}
-->
</HEAD>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }
</script>
<body>
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#DEDFDE" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="500">
  <tr align="Center" style="color:White;background-color:#6B696B;font-weight:bold;"> 	
    <td style="width:130px;" width="130"><font color="#FFFFFF">标题</font></td>
    <td width="299"><font color="#FFFFFF">所在省市</font></td>
    <td style="width:160px;" width="182"><font color="#FFFFFF">发表日期</font></td>
    <td style="width:30px;" width="125" align="center"><font color="#FFFFFF">详情</font></td>
  </tr>
  <%
    const MaxPerPage=20
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
        fz=" and (province='"&session("fz")&"' or city='"&session("fz")&"')"
		end if					
    sql="select  * from repair_info where 1=1"&fz&" order by infoid desc"
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
  <tr align="Center" style="color:#003399;background-color:White;" bgcolor="#F7F7DE"> 
    <td width="130"> 
      <%if len(rs("title"))>6 then%>
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
  <%end sub%>
</table>
</body>
</html>
