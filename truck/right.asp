<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
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
<%if request("action")="" then%>
<table cellspacing="0" cellpadding="2" rules="cols" width="500" bordercolor="#DEDFDE" border="1" align="center">
  <tr align="Center" style="color:White;background-color:#6B696B;font-weight:bold;"> 
    <td width="43" height="25"> 
      <div align="center"><font color="#FFFFFF">类</font></div>
    </td>
    <td width="86" height="25"> 
      <div align="center"><font color="#FFFFFF">出发地点</font></div>
    </td>
    <td width="66" height="25"> 
      <div align="center"><font color="#FFFFFF">目的地</font></div>
    </td>
    <td width="65" height="25"> 
      <div align="center"><font color="#FFFFFF">类型</font></div>
    </td>
    <td width="66" height="25"> 
      <div align="center"><font color="#FFFFFF">重量(吨)</font></div>
    </td>
    <td width="64" height="25"> 
      <div align="center"><font color="#FFFFFF">发布日期</font></div>
    </td>
    <td width="66" height="25"> 
      <div align="center"><font color="#FFFFFF">详情</font></div>
    </td>
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
				
    sql="select  * from info2 where infotype='车源信息' or infotype='货源信息' order by addtime desc"
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


	
sub showContent 
dim i
i=0
%>
  <%do while not rs.eof%>
  <tr align="Center" style="background-color:#F7F7DE;"> 
    <td> 
      <%if rs("infotype")="车源信息" then %>
      <div align="center"><img src="../images/car.gif" width="25" height="20"> 
      </div>
      <%else%>
      <div align="center"><img src="../images/goods.gif" width="25" height="20"> 
      </div>
      <%end if%>
    </td>
    <td align="Right"> 
      <div align="center"><%=rs("city")%>&nbsp;<%=rs("area")%>→</div>
    </td>
    <td align="Left"> 
      <div align="center"><%=rs("city2")%>&nbsp;<%=rs("area2")%></div>
    </td>
    <td> 
      <div align="center"><%=rs("cartype")%></div>
    </td>
    <td style="font-family:宋体;"> 
      <div align="center"><%=rs("carload")%>吨</div>
    </td>
    <td> 
      <div align="center"> <%=month(rs("addtime"))%>-<%=day(rs("addtime"))%>&nbsp; 
        <%=hour(rs("addtime"))%>:<%=minute(rs("addtime"))%> </div>
    </td>
    <td> 
      <div align="center"><A href="javascript:openwindow('detail.asp?InfoID=<%=rs("InfoID")%>',500,400)">详情</A></div>
    </td>
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
</table>
<%end if%>
<%end if%>
<%if request("action")="search" then%>

<table cellspacing="1" cellpadding="2" rules="cols" border="0"  align="center" width="500" bgcolor="#DEDFDE">
  <tr align="Center" style="color:White;background-color:#6B696B;font-weight:bold;"> 
    <td width="43" height="25"> 
      <div align="center"><font color="#FFFFFF">类</font></div>
    </td>
    <td width="87" height="25"> 
      <div align="center"><font color="#FFFFFF">出发地点</font></div>
    </td>
    <td width="77" height="25"> 
      <div align="center"><font color="#FFFFFF">目的地</font></div>
    </td>
    <td width="45" height="25"> 
      <div align="center"><font color="#FFFFFF">类型</font></div>
    </td>
    <td width="68" height="25"> 
      <div align="center"><font color="#FFFFFF">重量(吨)</font></div>
    </td>
    <td width="79" height="25"> 
      <div align="center"><font color="#FFFFFF">发布日期</font></div>
    </td>
    <td width="57" height="25"> 
      <div align="center"><font color="#FFFFFF">详情</font></div>
    </td>
  </tr>
  <%
     const MaxPerPage2=15
   	dim totalPut2   
   	dim CurrentPage2
   	dim TotalPages2

   	dim idlist2
   	if not isempty(request("page")) then
      		CurrentPage2=cint(request("page"))
   	else
      		CurrentPage2=1
   	end if
dim strUnit2,sfilename2
strUnit2="个信息"
dim sql2,city,city2,area,area2,province,province2
dim rs2,type_search
        type_search=request("type_search")
      select case type_search
	         case 1
        if request("city")<>"地级市" then
	      province=" and province='"&request("province")&"' and city='"&request("city")&"'"
		  else 
		  if request("province")<>"省份" then
	       province=" and province='"&request("province")&"'"
		  else
		   province=""
		  end if
	    end if
        if request("city2")<>"地级市" then
		  if request("province")="省份" then
	      province2=" and (province2='"&request("province2")&"' and city2='"&request("city2")&"')"
		  else
		  province2=" or (province2='"&request("province2")&"' and city2='"&request("city2")&"')"
		  end if
		  else 
		  if request("province")<>"省份" then
	       province2=" or province2='"&request("province2")&"'"
		  else
		   province=""
		  end if	    
		end if		
		
		   case 2
        if request("city")<>"地级市" then
	      province=" and province='"&request("province")&"' and city='"&request("city")&"'"
		  else 
		  if request("province")<>"省份" then
	       province=" and province='"&request("province")&"'"
		  else
		   province=""
		  end if
	    end if
        if request("city2")<>"地级市" then
	      province2=" and (province2='"&request("province2")&"' and city2='"&request("city2")&"')"
		  else 
		  if request("province")<>"省份" then
	       province2=" and province2='"&request("province2")&"'"
		  else
		   province=""
		  end if	    
		end if	
		
		case 3
        if request("city")<>"地级市" then
	      province=" and province='"&request("province")&"' and city='"&request("city")&"' and infotype='货源信息'"
		  else 
		  if request("province")<>"省份" then
	       province=" and province='"&request("province")&"' and infotype='货源信息'"
		  else
		   province="and infotype='货源信息'"
		  end if
	    end if
        if request("city2")<>"地级市" then
	      province2=" or (province2='"&request("province2")&"' and city2='"&request("city2")&"' and infotype='货源信息')"
		  else 
		  if request("province")<>"省份" then
	       province2=" or (province2='"&request("province2")&"' and infotype='货源信息')"
		  else
		   province=""
		  end if	    
		end if				   
	  case else 
        if request("city")<>"地级市" then
	      province=" and province='"&request("province")&"' and city='"&request("city")&"' and infotype='车源信息'"
		  else 
		  if request("province")<>"省份" then
	       province=" and province='"&request("province")&"' and infotype='车源信息'"
		  else
		   province="and infotype='车源信息'"
		  end if
	    end if
        if request("city2")<>"地级市" then
	      province2=" or (province2='"&request("province2")&"' and city2='"&request("city2")&"' and infotype='车源信息')"
		  else 
		  if request("province")<>"省份" then
	       province2=" or (province2='"&request("province2")&"' and infotype='车源信息')"
		  else
		   province=""
		  end if	    
		end if	
      end select 	
        sql2="select  * from info2 where 1=1"&province&province2&" order by addtime desc"
  sfilename2="right.asp"
	'sql2="select  * from file_info where 1=1"&city&city2&area&area2&" order by infoid desc"
    Set rs2= Server.CreateObject("ADODB.Recordset")
	rs2.open sql2,conn,1,1
  	if rs2.eof and rs2.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 信 息</td></tr>"
   	else
%>
  <%
      		totalPut2=rs2.recordcount
      		if CurrentPage2<1 then
          		CurrentPage2=1
      		end if
      		if (CurrentPage2-1)*MaxPerPage2>totalPut2 then
	   		if (totalPut2 mod MaxPerPage2)=0 then
	     			CurrentPage2= totalPut2 \ MaxPerPage2
	  		else
	      			CurrentPage2= totalPut2 \ MaxPerPage2 + 1
	   		end if

      		end if
       		if CurrentPage2=1 then
            		showContent2
            		call showpage2(sfilename2,totalPut2,MaxPerPage2,false,false,strUnit2)
       		else
          		if (CurrentPage2-1)*MaxPerPage2<totalPut2 then
            			rs2.move  (CurrentPage2-1)*MaxPerPage2
            			dim bookmark3
            			bookmark3=rs2.bookmark3
            			showContent2
             			call showpage2(sfilename2,totalPut2,MaxPerPage2,false,false,strUnit2)
        		else
	        		CurrentPage2=1
           			showContent2
           			call showpage2(sfilename2,totalPut2,MaxPerPage2,false,false,strUnit2)
	      		end if
	   	end if
 
	sub showContent2
	dim k
	k=0 
%>
  <%do while not rs2.eof%>
  <tr align="Center" style="background-color:#F7F7DE;"> 
    <td> 
      <%if rs2("infotype")="车源信息" then %>
      <div align="center"><img src="../images/car.gif" width="25" height="20"> 
      </div>
      <%else%>
      <div align="center"><img src="../images/goods.gif" width="25" height="20"> 
      </div>
      <%end if%>
    </td>
    <td align="Right"> 
      <div align="center"><%=rs2("city")%>&nbsp;<%=rs2("area")%>→</div>
    </td>
    <td align="Left"> 
      <div align="center"><%=rs2("city2")%>&nbsp;<%=rs2("area2")%> </div>
    </td>
    <td> 
      <div align="center"><%=rs2("cartype")%></div>
    </td>
    <td style="font-family:宋体;"> 
      <div align="center"><%=rs2("carload")%>吨</div>
    </td>
    <td> 
      <div align="center"><%=rs2("addtime")%></div>
    </td>
    <td> 
      <div align="center"><A href="javascript:openwindow('detail.asp?InfoID=<%=rs2("InfoID")%>',500,400)">详情</A></div>
    </td>
  </tr>
  <tr align="Center" style="background-color:#F7F7DE;" bgcolor="#F7F7DE"> 
    <td colspan="7"> 
      <% 
k=k+1
if k>=MaxPerPage2 then exit do 
rs2.movenext 
loop 
rs2.close
set rs2=nothing
%>
      <%end sub%>
      <%end if%>
      <%end if%>
    </td>
  </tr>

</table>

</BODY></HTML>