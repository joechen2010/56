
<META http-equiv="Content-Type" content="text/html; chars73et=gb2312">
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#3366CC" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="778">
    <tr align="Center" style="color:#CCCCFF;background-color:#003399;font-weight:bold;">		
      <td style="width:130px;" width="130">车牌号</td>
      <td width="149">车籍</td>
      <td width="150">车主</td>
      <td style="width:160px;" width="182">车型</td>
      <td width="125" align="center">详情</td>
	</tr>
                    <%
     dim rs7
    sql="select  * from file_info a,qyml b where b.user='"&user&"' and  a.gsid='"&user&"' order by infoid desc"
    Set rs7= Server.CreateObject("ADODB.Recordset")
	rs7.open sql,conn,1,1
  	if rs7.eof and rs7.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> 还 没 有 任 何 信 息</td></tr> </p>"
   	else

%>  

				<%do while not rs7.eof%>		
	
	<tr align="Center" style="color:#003399;background-color:White;">		
      <td width="130"><%=rs7("carnum")%></td>
      <td  width="149" align="center"><%=rs7("city")%></td>
      <td  width="150" align="center"><%=rs7("qymc")%></td>
      <td width="182" align="center"><%=rs7("cartype")%></td>
      <td width="125"><A href="javascript:openwindow('../cheliang/detail.asp?InfoID=<%=rs7("InfoID")%>',500,400)">详情</A></td>
	</tr>
<% 
rs7.movenext 
loop 
%>

<%end if%> 
</table>