<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<HTML>
<HEAD>
<title>配货信息查询结果</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css rel=stylesheet>
<STYLE type=text/css>.bg {
	BACKGROUND-POSITION: 50% top; BACKGROUND-REPEAT: no-repeat
}
</STYLE>
<META content="诚信物流网-配货信息查询结果" name="DESCRIPTION">
<META content="配货信息,车,货" name="KEYWORDS">
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
<meta name="vs_defaultClientScript" content="JavaScript">
</HEAD>
<body leftmargin="0" topmargin="0">
<embed title="中华物流网" src="../truck/images/banner.swf" width="500" height="60" type="application/x-shockwave-flash">
</embed>
<table width=98% border=0 align=center cellpadding=4 cellspacing=0>
  <tbody> 
  <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from zx_info a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
  <tr align="center" bgcolor="#B1EBC6"> 
    <td class=title02 colspan=4><b><font size="4"><%=rs("a.city")%>←→<%=rs("city2")%></font></b></td>
  </tr>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=4 height=1></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7 width="79">线路描述：</td>
    <td colspan="3" valign="top" width="431"><%=rs("content")%> </td>
  </tr>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=4 height=1></td>
  </tr>
  <tr> 
    <td colspan=4></td>
  </tr>
  </tbody> 
</table>
<br>
<table width="98%" border="0" cellspacing="1" cellpadding="5" align="center" bordercolor="#666666" bgcolor="#3366CC">
  <tr align="center"> 
    <td width="132"><font color="#FFFFFF">区间</font></td>
    <td width="49"><font color="#FFFFFF">公里</font></td>
    <td width="136"><font color="#FFFFFF">重量单价</font></td>
    <td width="138"><font color="#FFFFFF">体积单价</font></td>
  </tr>
  <%
						   set rs1=server.CreateObject("adodb.recordset")
						    sql1 = "select * from zx_info2 where infoid="&rs("infoid")&" order by id desc"
						   rs1.open sql1,conn,1,1
						   if not rs1.eof then
						    for j=1 to rs1.recordcount
						   %>
  <tr bgcolor="#FFFFFF" align="center"> 
    <td width="132"><%=rs("a.city")%>←→<%=rs1("city")%></td>
    <td width="49">&nbsp;</td>
    <td width="136"><%=rs1("prices")%>元/吨</td>
    <td width="138"><%=rs1("tiji")%>元/M3</td>
  </tr>
  <%
							  rs1.movenext
							  next
							  end if
							  %>
</table>
<br>
<table width=98% border=0 align=center cellpadding=4 cellspacing=1 bgcolor="#e7e7ff">
  <tbody> 
  <tr bgcolor="#408080"> 
    <td class=title02 colspan=2 align="center"><b><font color="#FFFFFF">联系方式</font></b></td>
  </tr>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=2 height=1></td>
  </tr>
  <tr> 
    <td width="151"><b>单位名称：</b></td>
    <td width="829"><%=rs("a.name")%></td>
  </tr>
  <tr> 
    <td width="151"><b>联系人：</b></td>
    <td width="829"><%=rs("b.name")%></td>
  </tr>
  <tr> 
    <td width="151"><b>联系电话：</b></td>
    <td width="829"><%=rs("phone2")%>-<%=rs("phone3")%></td>
  </tr>
  <tr> 
    <td width="151"><b>联系地址：</b></td>
    <td width="829"><%=rs("Address")%></td>
  </tr>
  </tbody> 
</table>
</body>
</HTML>
