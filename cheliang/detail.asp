<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
        <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from file_info a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
<HTML>
<HEAD>
<title>车辆档案--<%=rs("carnum")%></title>
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
<table width=480 border=0 align=center cellpadding=0 cellspacing=0>
  <tbody> 
  <tr> 
    <td height=20><embed src="../flash/banner.swf" width="480" height="60" type="application/x-shockwave-flash">
      </embed></td>
  </tr>
  <tr bgcolor="#B1EBC6" align="center"> 
    <td class=title02 height="20"><b>车辆档案--<%=rs("carnum")%></b></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7> 
      <table border=0 align=center cellpadding=2 cellspacing=2 width=100%>
        <tbody> 
        <tr> 
          <td width="16%" bgcolor=#f7f7f7><b>车牌号：</b></td>
          <td width="33%" align="center"><%=rs("carnum")%></td>
          <td width="16%" bgcolor=#f7f7f7><b>车主：</b></td>
          <td colspan="3" width="35%"><%=rs("qymc")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7 width="16%"><b>品牌：</b></td>
          <td align="center" width="33%"><%=rs("carpp")%></td>
          <td bgcolor=#f7f7f7 width="16%"><b>车籍：</b></td>
          <td colspan="3" width="35%"><%=rs("city")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7 width="16%"><b>车型：</b></td>
          <td align="center" width="33%"><%=rs("cartype")%></td>
          <td bgcolor=#f7f7f7 width="16%"><b>载重：</b></td>
          <td colspan="3" width="35%"><%=rs("carload")%>吨</td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7 width="16%"><b>车长：</b></td>
          <td align="center" width="33%"><%=rs("carlong")%>米</td>
          <td bgcolor=#f7f7f7 width="16%"><b>出厂时间：</b></td>
          <td colspan="3" width="35%"><%=rs("time")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7 width="16%"><b>主要司机：</b></td>
          <td align="center" width="33%"><%=rs("b.siji")%></td>
          <td bgcolor=#f7f7f7 width="16%"><b>该车状况：</b></td>
          <td colspan="3" width="35%"><%=rs("content")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7 width="16%"><b>发布时间：</b></td>
          <td colspan="5"><%=rs("addtime")%></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  <tr align="center"> 
    <td bgcolor=#EFEFEF height=27>168货运信息网对于本条信息确性的声明：</td>
  </tr>
  <tr> 
    <td height=54><font color="#999999">以上信息我网未做亲自核实，我网将力求但不能确保信息的准确性。您若发现此为虚假信息立即投诉，一经核实将删除其会员帐号！<br>
      声明：此条信息不得转载到其他网站，违者将被子停止帐号服务。</font></td>
  </tr>
  <tr align="center" bgcolor="#B1EBC6"> 
    <td height=20>[<a id="HyperLink1" target="_blank">查看该会员网站</a>] [<a href="JavaScript:window.print()">打印此页</a>] 
      [<a href="javascript:window.close()">关闭窗口</a>]</td>
  </tr>
  </tbody> 
</table>
</body>
</HTML>
