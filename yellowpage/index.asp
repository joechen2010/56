<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<html>
<head>
<title>������������������ҳ</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css 
rel=stylesheet>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
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
    <td width="66%" height="26" class="text"><font color="#FFFFFF">������ҳ</font></td>
    <td width="34%" height="24" align="right" class="text">&nbsp;</td>
  </tr>
</table>
<table align="center" width="778" cellpadding="0" cellspacing="0" border="0">
  <tr> 
    <td align="center" width="52%" valign="top"> 
      <table width="99%" border="0" cellspacing="1" cellpadding="0" bgcolor="#FF9900" align="left">
        <tr>
          <td bgcolor="#FFFFFF"> 
            <div align="center"><img src="images/1.gif" width="334" height="269" vspace="40" border="0" usemap="#Map"></div>
          </td>
        </tr>
      </table>
      
    </td>
    <td align="center" width="48%" valign="top">
	<table cellspacing="0" cellpadding="2" align="Center" bordercolor="Tan" border="0" id="DataGrid1" style="color:Black;background-color:LightGoldenrodYellow;border-color:Tan;border-width:1px;border-style:solid;font-size:X-Small;width:100%;border-collapse:collapse;">
	<tr align="Center" style="background-color:Tan;font-weight:bold;">
		<td align="center">����</td><td align="center">������</td><td align="center">��λ����</td>
	</tr>
    <%if request("action")="" then%>
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
strUnit="����Ϣ"
				
    sql="select * from qyml order by id desc"
	sfilename="index.asp"
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> �� û �� �� �� �� Ա</td></tr> </p>"
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
		<td><span class="STYLE1"><%=rs("qylb")%></span></td>
		<td><%=rs("city")%></td>
		<td>
				  <%if len(rs("qymc"))>10 then%>
				  <a href="../site/aboutus.asp?InfoID=<%=rs("ID")%>&user=<%=rs("user")%>" target="_blank">
				  
				  <%=left(rs("qymc"),8)%>...</a>
				  <%else%>
				  <a href="../site/aboutus.asp?InfoID=<%=rs("ID")%>&user=<%=rs("user")%>" target="_blank">
				  
				  <%=rs("qymc")%></a>
				  <%end if%>		
		</td>
	</tr>
<% 
i=i+1
if i>=MaxPerPage then exit do 
rs.movenext 
loop 

%>
<%end sub%> 
<%end if%> 
<%if request("action")="search" then%>
                    <%
dim sql2,city,infotype,kftype
dim rs2

        if request("province")<>"ʡ��" then
	    city=" and province='"&request("province")&"'"
	    end if
	    if request("infotype")<>"" then
	    infotype=" and infotype='"&request("infotype")&"'"
	    end if
	    if request("area")<>"" then
	    area=" and area='"&request("area")&"'"
	    end if
	    if request("qylb")<>"" then
	    qylb=" and qylb='"&request("qylb")&"'"
	    end if				
    sql2="select  * from qyml where 1=1"&city&qylb&" order by id desc"
	'sql="select  * from ProdMain  where 1=1"&sql1&sql2&sql3&sql4&" and larcode='������Ϣ' order by adddate desc"
    Set rs2= Server.CreateObject("ADODB.Recordset")
	rs2.open sql2,conn,1,1
  	if rs2.eof and rs2.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> �� û �� �� �� �� Ϣ</td></tr>"
   	else
%>  

				<%do while not rs2.eof%>
	<tr align="Center" onMouseOver="this.style.backgroundColor='#efefef'" onMouseOut="this.style.backgroundColor=''">
		<td><%=rs2("qylb")%></td>
		<td><%=rs2("city")%></td>
		<td><a href="detail.asp?infoid=<%=rs2("id")%>" target="_blank">
		<%=rs2("qymc")%></a></td>
	</tr>
<% 

rs2.movenext 
loop 

%>
<%end if%>

<%end if%>
	
</table>
</td>
  </tr>
</table>
<!--#include file="../inc/down.htm"-->
<map name="Map"> 
  <area shape="rect" coords="270,180,309,197" href="zj.asp?province=�㽭">
<area shape="rect" coords="272,30,327,51" href="hlj.asp?province=������">
<area shape="rect" coords="278,63,316,79" href="jl.asp?province=����">
<area shape="rect" coords="263,86,295,100" href="ln.asp?province=����">
<area shape="rect" coords="42,86,77,103" href="xj.asp?province=�½�">
<area shape="rect" coords="50,159,84,177" href="xz.asp?province=����">
<area shape="rect" coords="121,100,149,115" href="gs.asp?province=����">
<area shape="rect" coords="159,96,211,109" href="nmg.asp?province=���ɹ�">
<area shape="rect" coords="101,126,140,141" href="qh.asp?province=�ຣ">
<area shape="rect" coords="146,129,186,143" href="nx.asp?province=����">
<area shape="rect" coords="135,162,170,177" href="sc.asp?province=�Ĵ�">
<area shape="rect" coords="139,217,178,231" href="yn.asp?province=����">
<area shape="rect" coords="170,199,205,211" href="gz.asp?province=����">
<area shape="rect" alt="����" coords="174,245,207,261" href="hn.asp?province=����">
<area shape="rect" alt="����" coords="184,219,218,231" href="gx.asp?province=����">
<area shape="rect" alt="�㶫" coords="222,217,251,233" href="gd.asp?province=�㶫">
<area shape="rect" alt="����" coords="165,180,201,194" href="cq.asp?province=����">
<area shape="rect" alt="����" coords="210,199,239,212" href="hu_nan.asp?province=����">
<area shape="rect" alt="����" coords="252,199,288,212" href="hj.asp?province=����">
<area shape="rect" alt="����" coords="245,162,275,176" href="an_hui.asp?province=����">
<area shape="rect" alt="����" coords="230,93,259,106" href="beijing.asp?province=����">
<area shape="rect" alt="�ӱ�" coords="225,110,257,120" href="hb.asp?province=�ӱ�">
<area shape="rect" alt="ɽ��" coords="203,122,237,136" href="sx.asp?province=ɽ��">
<area shape="rect" alt="����" coords="182,144,214,159" href="shan_xi.asp?provinec=����">
<area shape="rect" alt="ɽ��" coords="240,124,276,138" href="sd.asp?province=ɽ��">
<area shape="rect" alt="����" coords="217,144,243,158" href="he_nan.asp?province=����">
<area shape="rect" alt="����" coords="256,141,290,158" href="js.asp?province=����">
<area shape="rect" alt="����" coords="209,165,239,180" href="hubei.asp?province=����">
<area shape="rect" alt="����" coords="235,181,269,196" href="jx.asp?province=����">
<area shape="rect" alt="���" coords="258,107,296,122" href="tianj.asp?province=���"><area shape="rect" coords="281,161,314,175" href="sh.asp?province=�Ϻ�">
  <area shape="rect" coords="291,217,325,231" href="tw.asp?province=%CC%A8%CD%E5">
  <area shape="rect" coords="256,229,287,241" href="xg.asp?province=%CF%E3%B8%DB">
  <area shape="rect" coords="221,237,248,250" href="am.asp?province=%B0%C4%C3%C5">
</map>
</body>
</html>