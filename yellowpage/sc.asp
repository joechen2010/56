
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<html>
<head>
<title>诚信物流网·物流黄页</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css 
rel=stylesheet>
</head>
<body>
<!--#include file="../inc/top.htm"-->
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="center"><iframe id="ifrsearch" src="../search/index_login.asp" width="100%" frameborder="0" scrolling="no" height="90"></iframe></td>
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
    <td width="66%" height="26" class="text"><font color="#FFFFFF">物流黄页</font></td>
    <td width="34%" height="24" align="right" class="text">&nbsp;</td>
  </tr>
</table>
<table align="center" width="778" cellpadding="0" cellspacing="0" border="0">
  <tr> 
    <td align="center" width="52%" valign="top"> 
      <table width="99%" border="0" cellspacing="1" cellpadding="0" bgcolor="#FF9900" align="left">
        <tr>
          <td bgcolor="#FFFFFF" align="center"><a href="index.asp">返回到全国地图</a></td>
        </tr>	  
        <tr>
          <td bgcolor="#FFFFFF" height="396">
<IMG height=360 src="images/127.gif" width=520 useMap=#Map border=0>		  
		  </td>
        </tr>
      </table>
      
    </td>
    <td align="center" width="48%" valign="top">
	<table cellspacing="0" cellpadding="2" align="Center" bordercolor="Tan" border="0" id="DataGrid1" style="color:Black;background-color:LightGoldenrodYellow;border-color:Tan;border-width:1px;border-style:solid;font-size:X-Small;width:100%;border-collapse:collapse;">
	<tr align="Center" style="background-color:Tan;font-weight:bold;">
		<td align="center">类型</td><td align="center">所在市</td><td align="center">单位名称</td>
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
	if request("province")<> "" then			
    sql="select * from qyml where province='"&request("province")&"' order by id desc"
	else
	 if request("city")<>"" then
    sql="select * from qyml where city='"&request("city")&"' order by id desc"
	  else
	sql="select * from qyml order by id desc"
	end if
	end if
	sfilename="index.asp"
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 会 员</td></tr> </p>"
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
	<tr align="Center" onMouseOver="this.style.backgroundColor='#efefef'" onMouseOut="this.style.backgroundColor=''">
		<td><%=rs("qylb")%></td>
		<td><%=rs("city")%></td>
		<td><a href="../yellowpage/site.asp?InfoID=<%=rs("ID")%>&user=<%=rs("user")%>" target="_blank"><%=rs("qymc")%></a></td>
	</tr>
<% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 

%>
<%end sub%> 


</table>
</td>
  </tr>
</table>
<!--#include file="../inc/down.htm"-->
<MAP name=Map>
<AREA shape=RECT alt=成都 coords=290,153,329,166 href="sc.asp?city=成都">
<AREA shape=RECT alt=攀枝花 coords=195,304,265,322 href="sc.asp?city=攀枝花">
<AREA shape=RECT alt=德阳 coords=289,136,333,150 href="sc.asp?city=德阳">
<AREA shape=RECT alt=广元 coords=329,83,375,101  href="sc.asp?city=广元">
<AREA shape=RECT alt=内江 coords=304,195,356,206 href="sc.asp?city=内江">
<AREA shape=RECT alt=南充 coords=349,129,393,146 href="sc.asp?city=南充">
<AREA shape=RECT alt=宜宾 coords=296,235,345,250 href="sc.asp?city=宜宾">
<AREA shape=RECT alt=达州 coords=396,120,442,133 href="sc.asp?city=达州">
<AREA shape=RECT alt=巴中 coords=375,98,420,113 href="sc.asp?city=巴中">
<AREA shape=RECT alt=阿坝州(马尔康) coords=190,85,288,102 href="sc.asp?city=阿坝州">
<AREA shape=RECT alt=凉山州(西昌) coords=185,266,274,282 href="sc.asp?city=凉山州">
<AREA shape=RECT alt=自贡 coords=304,207,348,222 href="sc.asp?city=自贡">
<AREA shape=RECT alt=泸州 coords=334,224,390,239 href="sc.asp?city=泸州">
<AREA shape=RECT alt=绵阳 coords=292,107,345,119 href="sc.asp?city=绵阳">
<AREA shape=RECT alt=遂宁 coords=333,155,370,168 href="sc.asp?city=遂宁">
<AREA shape=RECT alt=乐山 coords=254,213,307,224 href="sc.asp?city=乐山">
<AREA shape=RECT alt=眉山 coords=274,177,318,193 href="sc.asp?city=眉山">
<AREA shape=RECT alt=广安 coords=370,159,415,176 href="sc.asp?city=广安">
<AREA shape=RECT alt=雅安 coords=232,177,272,193 href="sc.asp?city=雅安">
<AREA shape=RECT alt=资阳 coords=319,171,364,187 href="sc.asp?city=资阳">
<AREA shape=RECT alt=甘孜州(康定) coords=145,146,231,161 href="sc.asp?city=甘孜州"></MAP>
</body>
</html>

