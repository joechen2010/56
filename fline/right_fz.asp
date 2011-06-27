<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<LINK href="images/page.css" type=text/css rel=stylesheet>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }

</script>
</head>

<body>
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#DEDFDE" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="500">
  <tr align="Center" style="color:White;background-color:#6B696B;font-weight:bold;"> 
    <td height="26" width="100"><strong><font 
                        color=#ffffff>专线名称</font></strong></td>
    <td><strong><font 
                        color=#ffffff>途经(中转)城市</font></strong></td>
    <td width="120"><strong><font 
                        color=#ffffff>公司名称</font></strong></td>
    <td width="60"><strong><font 
                        color=#ffffff>专线</font></strong></td>
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
sfilename="right_fz.asp"
        if session("fz")<>"" then
        fz=" and (a.province='"&session("fz")&"' or a.city='"&session("fz")&"')"
		end if	
    sql="select  * from zx_info a,qyml b where 1=1"&fz&" and infotype='货运专线' and a.gsid=b.user order by infoid desc"
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
  <tr onMouseOver="this.style.backgroundColor='#FAEFE0'" 
                    style="CURSOR: hand; BACKGROUND-COLOR: #f9f9f9" 
                    onMouseOut="style.backgroundColor='#F9F9F9'" align="center"> 
    <td height=25><%=rs("a.city")%>←→<%=rs("city2")%></td>
    <td height=25> 
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
    <td height=25> 
      <%=rs("qymc")%> 
    </td>
    <td height=25><a href="javascript:openwindow('detail.asp?InfoID=<%=rs("InfoID")%>',500,400)">详情</a></td>
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
