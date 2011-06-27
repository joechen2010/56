
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<HTML>
<HEAD>
<title>专线信息检索条件</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<STYLE type=text/css>.bg {
	BACKGROUND-POSITION: 50% top; BACKGROUND-REPEAT: no-repeat
}
</STYLE>
</HEAD>
<body leftmargin="0" rightmargin="0" style="FONT-SIZE: 10pt" bgColor="#FFFFFF" topmargin="0">
<!--#include file="../inc/top_beijing.htm" -->
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="center"><iframe id="ifrsearch" src="../search/index_job.asp" width="100%" frameborder="0" scrolling="no" height="90"></iframe></td>
  </tr>
</table>
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<%if request("action")="" then%>
<table width="778" border="0" cellpadding="0" cellspacing="1" bgcolor="#FF6600" align="center">
  <tr bgcolor="FFC390">
    <td height="2" colspan="4"></td>
  </tr>
  <tr bgcolor="#FF6600">
    <td height="26" class="text"><font color="#FFFFFF">最新招聘信息</font></td>
    <td height="24" align="right" class="text"><a href="index_z_beijing.asp"><font color="#FFFFFF">more</font></a></td>
  </tr>
  <tr>
    <td height="26" colspan="2" class="text">

<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#3366CC" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="778">
    <tr align="Center" style="color:#CCCCFF;background-color:#003399;font-weight:bold;">		
      <td style="width:130px;" width="129">招聘职位</td>
      <td width="208">招聘单位</td>
      <td style="width:160px;" width="179">所在省市</td>
      <td style="width:160px;" width="157">发布日期</td>
      <td style="width:30px;" width="53" align="center">详情</td>
	</tr>
		  
                    <%
    sql1="select top 10 * from zp_info a,qyml b where infotype='招聘信息' and a.gsid=b.user and a.city='北京' order by infoid desc"
    Set rs1= Server.CreateObject("ADODB.Recordset")
	rs1.open sql1,conn,1,1
  	if rs1.eof and rs1.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 信 息</td></tr>"
   	else
%>  

				<%do while not rs1.eof%>		
	
	<tr align="Center" style="color:#003399;background-color:White;">		
      <td width="129"><%=rs1("job")%></td>
      <td  width="208" align="center"><%=rs1("qymc")%></td>
      <td width="179" align="center"><%=rs1("a.province")%>&nbsp;<%=rs1("a.city")%></td>
      <td width="157" align="center"><%=rs1("addtime")%></td>
      <td width="53"><A href="detail.asp?InfoID=<%=rs1("InfoID")%>">详情</A></td>
	</tr>
<% 
rs1.movenext 
loop 
%>
<%end if%> 
    </table>	
   </td>
  </tr>
  
  <tr bgcolor="#FF6600"> 
    <td width="66%" height="26" class="text"><font color="#FFFFFF">最新求职信息</font></td>
    <td width="34%" height="24" align="right" class="text"><a href="index_q_beijing.asp"><font color="#FFFFFF">more</font></a></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#3366CC" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="778">
    <tr align="Center" style="color:#CCCCFF;background-color:#003399;font-weight:bold;">		
      <td style="width:130px;" width="129">求职人</td>
      <td width="215">求职意向</td>
      <td style="width:160px;" width="180">所在省市</td>
      <td style="width:160px;" width="150">发布日期</td>
      <td style="width:30px;" width="52" align="center">详情</td>
	</tr>

                    <%
    sql="select top 10 * from zp_info a,qyml b where infotype='求职信息' and a.gsid=b.user and a.city='北京' order by infoid desc"
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 信 息</td></tr>"
   	else
dim i
i=0
%>  

				<%do while not rs.eof%>		
	
	<tr align="Center" style="color:#003399;background-color:White;">		
      <td width="129"><%=rs("name")%></td>
      <td  width="215" align="center"><%=rs("job")%></td>
      <td width="180" align="center"><%=rs("a.province")%>&nbsp;<%=rs("a.city")%></td>
      <td width="150" align="center"><%=rs("addtime")%></td>
      <td width="52"><A href="detail.asp?InfoID=<%=rs("InfoID")%>">详情</A></td>
	</tr>
                    <% 
i=i+1
rs.movenext 
loop 
%>
<%end if%> 
<%end if%> 
</table>
</td>
</tr>
</table>
<%if request("action")="search" then%>
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#3366CC" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="778">
    <tr align="Center" style="color:#CCCCFF;background-color:#003399;font-weight:bold;">		
      <td style="width:130px;" width="130">类型</td>
      <td width="299">所在省市</td>
      <td style="width:160px;" width="182">联系方式</td>
      <td style="width:30px;" width="125" align="center">详情</td>
	</tr>



                    <%
dim sql2,province,jobtype,area,area2
dim rs2

        if request("job")<>"" then
	    job=" and job='"&request("job")&"'"
	    end if	
        if request("jobtype")<>"" then
	    jobtype=" and jobtype='"&request("jobtype")&"'"
	    end if	
		 if request("infotype")<>"" then
	    infotype=" and infotype='"&request("infotype")&"'"
	    end if		   
		 if request("city")<>"" then
	    city=" and city='"&request("city")&"'"
	    end if	
	'city=trim(request.form("city"))
    'city2=trim(request.form("city2"))
    'area=trim(request("area"))
    'area2=trim(request("area2"))			
    sql2="select  * from zp_info where 1=1"&city&jobtype&infotype&city&"  order by infoid desc"
	'sql="select  * from ProdMain  where 1=1"&sql1&sql2&sql3&sql4&" and larcode='车辆信息' order by adddate desc"
    Set rs2= Server.CreateObject("ADODB.Recordset")
	rs2.open sql2,conn,1,1
  	if rs2.eof and rs2.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 信 息</td></tr>"
   	else
%>  

				<%do while not rs2.eof%>
	<tr align="Center" style="color:#003399;background-color:White;">		
      <td width="130"><%=rs2("infotype")%></td>
      <td  width="299" align="center"><%=rs2("province")%>&nbsp;<%=rs2("city")%></td>
      <td width="182" align="center"><%=rs2("contact")%></td>
      <td width="125"><A href="detail.asp?InfoID=<%=rs2("InfoID")%>">详情</A></td>
	</tr>
<% 

rs2.movenext 
loop 

%>
<%end if%>

<%end if%>	
</table>
	
	
	<!--#include file="../inc/down.htm"-->
</body>
</HTML>
