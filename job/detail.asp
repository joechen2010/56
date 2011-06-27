<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
        'dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from zp_info a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
<HTML>
<HEAD>
<title>招聘求职信息--<%=rs("job")%></title>
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
<table width=480 border=0 align=center cellpadding=0 cellspacing=0>
  <tbody> 
  <tr> 
    <td height=20><embed src="../flash/banner.swf" width="480" height="60" type="application/x-shockwave-flash">
      </embed></td>
  </tr>
  <tr bgcolor="#B1EBC6" align="center"> 
    <td class=title02 height="20"><b>招聘求职信息</b></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7>
      <table border=0 align=center cellpadding=2 cellspacing=2 width=100%>
        <tbody> 

        <tr> 
          <td width="20%" bgcolor=#f7f7f7><b>所在省市：</b></td>
          <td><%=rs("a.province")%> <%=rs("a.city")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7><b>类别：</b></td>
          <td><%=rs("infotype")%> </td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7> <b> 
            <%if rs("infotype")="招聘信息" then%>
            招聘单位： 
            <%else%>
            姓名： 
            <%end if%>
            </b></td>
          <td> 
            <%if rs("infotype")="招聘信息" then%>
            <%=rs("qymc")%> 
            <%else%>
            <%=rs("name")%> 
            <%end if%>
          </td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7><b>招聘职位：</b></td>
          <td><%=rs("job")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7><b>月薪：</b></td>
          <td><%=rs("money")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7><b>有效期：</b></td>
          <td><%=rs("validate")%>天</td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7><b>职位说明：</b></td>
          <td><%=rs("content")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7><b>联系人：</b></td>
          <td><%=rs("name")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7><b>电话：</b></td>
          <td><%=rs("phone2")%>-<%=rs("phone3")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7><b>发布时间：</b></td>
          <td><%=rs("addtime")%></td>
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
