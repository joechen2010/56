
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
          <td bgcolor="#FFFFFF" height="396"><img src="images/117.gif"  border="0" usemap="#Map"></td>
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
<map name="Map">
<area shape="rect" coords="334,199,398,217" href="jl.asp?city=延边朝鲜族自治州">
<area shape="rect" coords="210,137,255,154" href="jl.asp?city=长春">
<area shape="rect" coords="259,160,297,175" href="jl.asp?city=吉林">
<area shape="rect" coords="173,87,217,106" href="jl.asp?city=松原">
<area shape="rect" coords="199,210,244,227" href="jl.asp?city=辽源">
<area shape="rect" coords="226,236,265,249" href="jl.asp?city=梅河口">
<area shape="rect" coords="270,257,313,273" href="jl.asp?city=白山">
<area shape="rect" coords="224,271,265,286" href="jl.asp?city=通化">
<area shape="rect" coords="110,66,163,82" href="jl.asp?city=白城">
<area shape="rect" coords="170,183,216,197" href="jl.asp?city=四平">
</map>

</body>
</html>

