<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<HTML>
<HEAD>
<title>�����Ϣ��ѯ���</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css 
rel=stylesheet>
<STYLE type=text/css>.bg {
	BACKGROUND-POSITION: 50% top; BACKGROUND-REPEAT: no-repeat
}
</STYLE>
<META content="����������-�����Ϣ��ѯ���" name="DESCRIPTION">
<META content="�����Ϣ,��,��" name="KEYWORDS">
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
<meta name="vs_defaultClientScript" content="JavaScript">
</HEAD>
<body leftmargin="0" topmargin="0">
<embed title="�л�������" src="../truck/images/banner.swf" width="500" height="60" type="application/x-shockwave-flash">
</embed> 
<table width=96% border=0 align=center cellpadding=5 cellspacing=1>
  <tbody> 
  <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from zx_info a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
  <tr bgcolor="#B1EBC6" align="center"> 
    <td class=title02 colspan=2><b><font size="4"><%=rs("a.city")%>����<%=rs("city2")%></font></b></td>
  </tr>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=2 height=1></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7 width="120"><b>��·������</b></td>
    <td><%=rs("content")%></td>
  </tr>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=2 height=1></td>
  </tr>
  </tbody> 
</table>
<table width="96%" border="0" align="center" cellspacing="0" cellpadding="0" height="27">
  <tr> 
    <td align="center">����(;��)���м��˼�</td>
  </tr>
</table>
<table cellspacing="1" cellpadding="5" bordercolor="White" border="0" id="DataGrid1" style="background-color:White;border-color:White;border-width:2px;border-style:Ridge;font-size:X-Small;width:96%;" align="center">
  <tr align="Center" style="color:#E7E7FF;background-color:#4A3C8C;font-weight:bold;">
    <td><font color="#FFFFFF">����</font></td>
    <td><font color="#FFFFFF"></font></td>
    <td><font color="#FFFFFF">��������</font></td>
    <td><font color="#FFFFFF">�������</font></td>
  </tr>
  <%
						   set rs1=server.CreateObject("adodb.recordset")
						    sql1 = "select * from zx_info2 where infoid="&rs("infoid")&" order by id desc"
						   rs1.open sql1,conn,1,1
						   if not rs1.eof then
						    for j=1 to rs1.recordcount
						   %>
  <tr align="center"> 
    <td bgcolor="#FFFFFF"><%=rs("a.city")%>����<%=rs1("city")%></td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#FFFFFF"><%=rs1("prices")%>Ԫ/��</td>
    <td bgcolor="#FFFFFF"><%=rs1("tiji")%>Ԫ/M3</td>
  </tr>
  <%
							  rs1.movenext
							  next
							  end if
							  %>
</table>
<table width="96%" border="0" align="center" cellspacing="0" cellpadding="0" height="27">
  <tr>
    <td>ע��������Ϊ����ֵ�������ο�</td>
  </tr>
</table>
<table width=96% border=0 align=center cellpadding=5 cellspacing=1 bgcolor="#e7e7ff">
  <tbody> 
  <tr> 
    <td colspan="2" align="center" bgcolor="#408080"><b><font color="#FFFFFF">��ϵ��ʽ</font></b></td>
  </tr>
  <tr> 
    <td width="120"><b>��ҵ���ƣ�</b></td>
    <td><%=rs("a.name")%></td>
  </tr>
  <tr> 
    <td><b>��ϵ�ˣ�</b></td>
    <td><%=rs("b.name")%></td>
  </tr>
  <tr> 
    <td><b>��ϵ�绰��</b></td>
    <td><%=rs("phone2")%>-<%=rs("phone3")%></td>
  </tr>
  <tr> 
    <td><b>��ϵ��ַ��</b></td>
    <td><%=rs("Address")%></td>
  </tr>
  </tbody> 
</table>
</body>
</HTML>
