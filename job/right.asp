<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }

</script>
</head>

<body>
<table width="500" border="0" cellpadding="0" cellspacing="1" bgcolor="#FF6600" align="center">
  <tr bgcolor="FFC390">
    <td height="2" colspan="4"></td>
  </tr>
  <tr bgcolor="#FF6600">
    <td height="26" class="text"><font color="#FFFFFF">������Ƹ��Ϣ</font></td>
    <td height="24" align="right" class="text"><a href="index_z.asp?infotype=��Ƹ��Ϣ" target="right"><font color="#FFFFFF">more</font></a></td>
  </tr>
  <tr>
    <td height="26" colspan="2" class="text">

<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#3366CC" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="500">
    <tr align="Center" style="color:#CCCCFF;background-color:#003399;font-weight:bold;">		
      <td style="width:130px;" width="129">��Ƹְλ</td>
      <td width="208">��Ƹ��λ</td>
      <td style="width:160px;" width="179">����ʡ��</td>
      <td style="width:160px;" width="157">��������</td>
      <td style="width:30px;" width="53" align="center">����</td>
	</tr>
<%
    sql1="select top 10 * from zp_info a,qyml b where infotype='��Ƹ��Ϣ' and a.gsid=b.user order by infoid desc"
    Set rs1= Server.CreateObject("ADODB.Recordset")
	rs1.open sql1,conn,1,1
  	if rs1.eof and rs1.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> �� û �� �� �� �� Ϣ</td></tr>"
   	else
%>  

				<%do while not rs1.eof%>		
	
	<tr align="Center" style="color:#003399;background-color:White;">		
      <td width="129"><%=rs1("job")%></td>
      <td  width="208" align="center">
				  <%if len(rs1("qymc"))>8 then%>
				  <%=left(rs1("qymc"),6)%>...
				  <%else%>
				 <%=rs1("qymc")%>
				  <%end if%>	  
	  </td>
      <td width="179" align="center"><%=rs1("a.province")%>&nbsp;<%=rs1("a.city")%></td>
      <td width="157" align="center"><%=rs1("addtime")%></td>
      <td width="53"><A href="javascript:openwindow('detail.asp?InfoID=<%=rs1("InfoID")%>',500,400)">����</A></td>
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
    <td width="66%" height="26" class="text"><font color="#FFFFFF">������ְ��Ϣ</font></td>
    <td width="34%" height="24" align="right" class="text"><a href="index_z.asp?infotype=��ְ��Ϣ"><font color="#FFFFFF">more</font></a></td>
  </tr>
</table>
<table width="500" border="0" align="center" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
<table cellspacing="0" cellpadding="4" align="Center" rules="all" bordercolor="#3366CC" border="1" id="DataGrid1" style="background-color:White;border-color:#3366CC;border-width:1px;border-style:None;font-size:X-Small;border-collapse:collapse;" width="500">
    <tr align="Center" style="color:#CCCCFF;background-color:#003399;font-weight:bold;">		
      <td style="width:130px;" width="129">��ְ��</td>
      <td width="215">��ְ����</td>
      <td style="width:160px;" width="180">����ʡ��</td>
      <td style="width:160px;" width="150">��������</td>
      <td style="width:30px;" width="52" align="center">����</td>
	</tr>
<%
    sql="select top 10 * from zp_info a,qyml b where infotype='��ְ��Ϣ' and a.gsid=b.user order by infoid desc"
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
  	if rs.eof and rs.bof then
       		response.write "<tr align=""center"" bgcolor=""#FFDBBF""><td colspan=""7"" bgcolor=""#FFCEAA""> �� û �� �� �� �� Ϣ</td></tr>"
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
      <td width="52"><A href="javascript:openwindow('detail.asp?InfoID=<%=rs("InfoID")%>',500,400)">����</A></td>
	</tr>
                    <% 
i=i+1
rs.movenext 
loop 
%>
<%end if%> 
</table>
</td>
</tr>
</table>