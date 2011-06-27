<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<HTML>
<HEAD>
<title>供求信息</title>
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
<table width=480 border=0 align=center cellpadding=2 cellspacing=2>
  <tbody> 
  <tr> 
    <td colspan=6 height=2><embed src="../truck/images/banner.swf" width="480" height="60" type="application/x-shockwave-flash">
      </embed></td>
  </tr>
  <tr align="center"> 
    <td class=title02 height="20" colspan=6 bgcolor=#B1EBC6><b>供求信息</b></td>
  </tr>
  <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from gq_info a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
  <tr> 
    <td width="16%"><b>标题：</b></td>
    <td colspan="5"><%=rs("title")%></td>
  </tr>
  <tr> 
    <td><b>信息类型：</b></td>
    <td colspan="5"><%=rs("infotype")%></td>
  </tr>
  <tr> 
    <td><b>所在省市：</b></td>
    <td colspan="5"><%=rs("a.province")%>&nbsp;&nbsp;<%=rs("a.city")%>&nbsp;&nbsp;<%=rs("a.area")%></td>
  </tr>
  <tr> 
    <td><b>有效期：</b></td>
    <td colspan="5"><%=rs("validate")%>天</td>
  </tr>
  <tr> 
    <td><b>发布时间：</b></td>
    <td colspan="5"><%=rs("addtime")%></td>
  </tr>
  <tr> 
    <td><b>联系人：</b></td>
    <td colspan="5"><%=rs("name")%></td>
  </tr>
  <tr> 
    <td><b>电话：</b></td>
    <td colspan="5"><%=rs("phone2")%>-<%=rs("phone3")%></td>
  </tr>
  <tr> 
    <td><b>内容：</b></td>
    <td colspan="5"><%=rs("content")%></td>
  </tr>
  <tr align="center"> 
    <td colspan=6 bgcolor="#e7e7e7" height="20">168货运信息网对于本条信息确性的声明：</td>
  </tr>
  <tr> 
    <td colspan=6><font color="#999999">以上信息我网未做亲自核实，我网将力求但不能确保信息的准确性。您若发现此为虚假信息立即投诉，一经核实将删除其会员帐号！<br>
      声明：此条信息不得转载到其他网站，违者将被子停止帐号服务。 </font></td>
  </tr>
  <tr> 
    <td colspan=6 height="20" bgcolor="#B1EBC6">[<a id="HyperLink1" target="_blank">查看该会员网站</a>] 
      [<a href="JavaScript:window.print()">打印此页</a>] [<a href="javascript:window.close()">关闭窗口</a>]</td>
  </tr>
  </tbody> 
</table>
</body>
</HTML>
