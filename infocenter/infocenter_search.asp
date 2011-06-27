

<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
type_search=request("type_search")
province=request("province")
province2=request("province2")
city=request("city")
city2=request("city2")
%>
<HTML>
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language=JavaScript src="../inc/p_c_a2.js"></script>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }

</script>
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
<BODY leftMargin=0 topMargin=0>

<table cellspacing="0" cellpadding="2" rules="cols" bordercolor="#DEDFDE" border="1"  align="center" width="100%">
  <tr align="Center" style="color:White;background-color:#6B696B;font-weight:bold;"> 
    <td height="25"> 
      <div align="center">类</div>    </td>
    <td height="25"> 
      <div align="center">出发地点</div>    </td>
    <td height="25"> 
      <div align="center">目的地</div>    </td>
    <td height="25"> 
      <div align="center">类型</div>    </td>
    <td height="25"> 
      <div align="center">重量(吨)</div>    </td>
    <td height="25"> 
      <div align="center">发布日期</div>    </td>
    <td height="25"> 
      <div align="center">详情</div>    </td>
  </tr>
 <%
     const MaxPerPage=15
   	dim totalPut   
   	dim CurrentPage
   	dim TotalPages
   	dim j
   	dim idlist
   	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if
dim strUnit,sfilename
strUnit="个信息"	
dim sql2,city,city2,area,area2,province,province2
dim rs2,type_search
        type_search=request("type_search")
	if 	request("type_search")<>"" then
      select case type_search
	         case 1
        if request("city")<>"地级市" then
	      p=" and province='"&request("province")&"' and city='"&request("city")&"'"
		  else 
		  if request("province")<>"省份" then
	       p=" and province='"&request("province")&"'"
		  else
		   p=""
		  end if
	    end if
        if request("city2")<>"地级市" then
		  if request("province")="省份" then
	      p2=" and (province2='"&request("province2")&"' and city2='"&request("city2")&"')"
		  else
		  p2=" or (province2='"&request("province2")&"' and city2='"&request("city2")&"')"
		  end if
		  else 
		  if request("province")<>"省份" then
	       p2=" or province2='"&request("province2")&"'"
		  else
		   p2=""
		  end if	    
		end if		
		
		   case 2
        if request("city")<>"地级市" then
	      p=" and province='"&request("province")&"' and city='"&request("city")&"'"
		  else 
		  if request("province")<>"省份" then
	       p=" and province='"&request("province")&"'"
		  else
		   p=""
		  end if
	    end if
        if request("city2")<>"地级市" then
	      p2=" and (province2='"&request("province2")&"' and city2='"&request("city2")&"')"
		  else 
		  if request("province")<>"省份" then
	       p2=" and province2='"&request("province2")&"'"
		  else
		   p2=""
		  end if	    
		end if	
		
		case 3
        if request("city")<>"地级市" then
	      p=" and province='"&request("province")&"' and city='"&request("city")&"' and infotype='货源信息'"
		  else 
		  if request("province")<>"省份" then
	       p=" and province='"&request("province")&"' and infotype='货源信息'"
		  else
		   p="and infotype='货源信息'"
		  end if
	    end if
        if request("city2")<>"地级市" then
	      p2=" or (province2='"&request("province2")&"' and city2='"&request("city2")&"' and infotype='货源信息')"
		  else 
		  if request("province")<>"省份" then
	       p2=" or (province2='"&request("province2")&"' and infotype='货源信息')"
		  else
		   p2=""
		  end if	    
		end if				   
	  case else 
        if request("city")<>"地级市" then
	      p=" and province='"&request("province")&"' and city='"&request("city")&"' and infotype='车源信息'"
		  else 
		  if request("province")<>"省份" then
	       p=" and province='"&request("province")&"' and infotype='车源信息'"
		  else
		   p="and infotype='车源信息'"
		  end if
	    end if
        if request("city2")<>"地级市" then
	      p2=" or (province2='"&request("province2")&"' and city2='"&request("city2")&"' and infotype='车源信息')"
		  else 
		  if request("province")<>"省份" then
	       p2=" or (province2='"&request("province2")&"' and infotype='车源信息')"
		  else
		   p2=""
		  end if	    
		end if	
      end select 	
        sql2="select  * from info2 where 1=1"&p&p2&" order by addtime desc"
		else
		sql2="select  * from info2 order by addtime desc"
		end if
		sfilename="infocenter_search.asp?province="&request("province")&"&province2="&request("province2")&"&city="&request("city")&"&city2="&request("city2")&"&type_search="&request("type_search")&""
	'sql2="select  * from file_info where 1=1"&city&city2&area&area2&" order by infoid desc"
    Set rs2= Server.CreateObject("ADODB.Recordset")
	rs2.open sql2,conn,1,1
  	if rs2.eof and rs2.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 信 息</td></tr>"
   	else
      		totalPut=rs2.recordcount
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
            			rs2.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs2.bookmark
            			showContent
             			call showpage2(sfilename,totalPut,maxperpage,false,false,strUnit)
        		else
	        		currentPage=1
           			showContent
           			call showpage2(sfilename,totalPut,maxperpage,false,false,strUnit)
	      		end if
	   	end if


	
sub showContent 
dim j
j=0	
%>  


				<%do while not rs2.eof%>
  <tr align="Center" style="background-color:#F7F7DE;"> 
    <td> 
	<%if rs2("infotype")="车源信息" then %>
      <div align="center"><img src="../images/car.gif" width="25" height="20">      </div>
	  <%else%>
	        <div align="center"><img src="../images/goods.gif" width="25" height="20">      </div>
	  <%end if%>    </td>
    <td align="Right"> 
      <div align="center"><%=rs2("city")%>&nbsp;<%=rs2("area")%>→</div>    </td>
    <td align="Left"> 
      <div align="center"><%=rs2("city2")%>&nbsp;<%=rs2("area2")%> </div>    </td>
    <td> 
      <div align="center"><%=rs2("cartype")%></div>    </td>
    <td style="font-family:宋体;"> 
      <div align="center"><%=rs2("carload")%>吨</div>    </td>
    <td> 
      <div align="center"><%=rs2("addtime")%></div>    </td>
    <td> 
      <div align="center"><A href="javascript:openwindow('../truck/detail.asp?InfoID=<%=rs2("InfoID")%>',800,400)">详情</A></div>    </td>
  </tr>
  
<% 
j=j+1
if j>=MaxPerPage then exit do 
rs2.movenext 
loop 
rs2.close
set rs2=nothing
%>
<%end sub%>
<%end if%>

</table>

</BODY></HTML>