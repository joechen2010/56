
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
set hit=server.CreateObject ("ADODB.RecordSet")
sql="select * from info2 where infoid="&request("infoid")
hit.Open sql,conn,1,3
if hit("hits")="" then
	hit("hits")=1
else
	hit("hits")=hit("hits")+1
end if
hit.update
hit.close
set hit=nothing
%>
<HTML>
<HEAD>
<title>配货信息查询结果</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css 
rel=stylesheet>
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
<table width=500 border=0 align=center cellpadding=2 cellspacing=2>
  <tbody> 
  <tr> 
    <td colspan=2 height=2><embed title="中华物流网" src="images/banner.swf" width="500" height="60" type="application/x-shockwave-flash">
      </embed></td>
  </tr>
  <tr> 
    <td class=title02 colspan=2 height="20" bgcolor="#B1EBC6" align="center"><b>货物信息</b></td>
  </tr>
  <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from info2 a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		  if not (rs.eof and rs.bof) then
		%>
  <tr> 
    <td width="16%" bgcolor=#f7f7f7><b>发车地点：</b></td>
    <td><%=rs("a.province")%>&nbsp;&nbsp;<%=rs("a.city")%>&nbsp;&nbsp;<%=rs("a.area")%></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>到达地点：</b></td>
    <td><%=rs("province2")%>&nbsp;&nbsp;<%=rs("city2")%>&nbsp;&nbsp;<%=rs("area2")%></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>车辆类型：</b></td>
    <td><%=rs("cartype")%> <b>载重：</b><%=rs("carload")%>吨 <b>车长：</b><%=rs("carlong")%>米</td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>有效期：</b></td>
    <td><%=rs("validate")%>天</td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>备注：</b></td>
    <td><%=rs("content")%></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>联系人：</b></td>
    <td> 
      <%if session("user")<>"" then%>
      <%=rs("qymc")%>&nbsp;&nbsp;<%=rs("name")%> 
      <%else
		    response.Write "<font color='#FF0000'>您还没有登陆，联系人不显示</font>"
			end if
		  %>    </td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>电话：</b></td>
    <td> 
      <%if session("user")<>"" then%>
      <%=rs("phone2")%>-<%=rs("phone3")%> 
      <%else
		    response.Write "<font color='#FF0000'>只有会员登录后才能查看</font>"
			end if
		  %>
      <%if session("user")="" then%>
      <% Response.Write(" <a href='../Reg/User_Reg.asp' target='_blank'><font color='#0000FF'>还不是会员请注册</font>")
    Response.Write(" <a href='../Login/login.asp' target='_blank'><font color='#0000FF'>是会员请登录</font></a>")
	%>
      <%end if%>
    &nbsp;&nbsp;<%if session("user")<>"" then%>
      <%=rs("mobile")%> 
      <%else
		    response.Write "您还没有登陆，电话不显示"
			end if
		  %></td>
  </tr>
  
  <tr> 
    <td bgcolor=#f7f7f7><b>发布时间：</b></td>
    <td><%=rs("addtime")%></td>
  </tr>
  <%end if%>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=2 height=1></td>
  </tr>
  <tr> 
    <td colspan=2 height=20 align="center" bgcolor="#EFEFEF">168货运信息网对于本条信息确性的声明：</td>
  </tr>
  <tr> 
    <td colspan=2><font color="#999999">以上信息我网未做亲自核实，我网将力求但不能确保信息的准确性。您若发现此为虚假信息立即投诉，一经核实将删除其会员帐号！<br>
      声明：此条信息不得转载到其他网站，违者将被子停止帐号服务。 </font></td>
  </tr>
  <tr bgcolor="#B1EBC6" align="center"> 
    <td colspan=2 height="20">[<a id="HyperLink1" target="_blank">查看该会员网站</a>] 
      [<a href="JavaScript:window.print()">打印此页</a>] [<a href="javascript:window.close()">关闭窗口</a>]</td>
  </tr>
  </tbody> 
</table>
</body>
</HTML>
