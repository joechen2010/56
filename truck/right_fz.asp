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
				
        if session("fz")<>"" then
        fz=" and (province='"&session("fz")&"' or city='"&session("fz")&"')"
		end if	
        sql="select  * from info2 where 1=1"&fz&" order by addtime desc"
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
end sub 
%>
</table>
<%end if%>
</BODY></HTML>