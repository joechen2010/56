
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
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
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>

<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <TBODY>
  <TR>
    <TD align="center" vAlign=top> 
      <TABLE width=100% border=0 align=center cellPadding=5 cellSpacing=5>
        <TBODY> 
        <TR> 
          <TD bgColor=#ff811e colSpan=4 height=2></TD>
        </TR>
        <TR> 
          <TD class=title02 bgColor=#ffeee0 colSpan=4>搬运吊装</TD>
        </TR>
        <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from banyun_info a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		  if not (rs.eof and rs.bof) then
		%>
        <TR> 
          <TD width="16%" bgColor=#f7f7f7>标题：</TD>
          <TD colspan="3"><%=rs("title")%></TD>
        </TR>
        <TR> 
          <TD bgColor=#e7e7e7 colSpan=4 height=1></TD>
        </TR>
        <TR> 
          <TD bgColor=#f7f7f7 width="16%">所在省市：</TD>
          <TD colspan="3"><%=rs("a.province")%> <%=rs("a.city")%> <%=rs("a.area")%></TD>
        </TR>
        <TR> 
          <TD bgColor=#e7e7e7 colSpan=4 height=1></TD>
        </TR>
        <TR> 
          <TD bgColor=#f7f7f7 width="16%">有效期：</TD>
          <TD width="22%"><%=rs("validate")%>天</TD>
          <TD bgColor=#f7f7f7 width="12%">发布时间：</TD>
          <TD width="50%"><%=rs("addtime")%></TD>
        </TR>
        <TR> 
          <TD bgColor=#f7f7f7 width="16%">联系人：</TD>
          <TD width="22%"><%=rs("name")%></TD>
          <TD bgColor=#f7f7f7 width="12%">电话：</TD>
          <TD width="50%"><%=rs("phone2")%>-<%=rs("phone3")%></TD>
        </TR>
        <TR> 
          <TD bgColor=#f7f7f7 width="16%">内容：</TD>
          <TD colspan="3"><%=rs("content")%></TD>
        </TR>
		<%end if%>
        <TR> 
          <TD bgColor=#e7e7e7 colSpan=4 height=1></TD>
        </TR>
        <TR> 
          <TD bgColor=#e7e7e7 colSpan=4 height=1></TD>
        </TR>
        <TR> 
          <TD bgColor=#ff811e colSpan=4 height=2></TD>
        </TR>
        </TBODY> 
      </TABLE>
      <BR>
      <TABLE cellSpacing=0 cellPadding=5 width=96% align=center border=0>
        <TBODY>
        <TR>
          <TD>
            <P><STRONG>168货运信息网对于本条信息确性的声明：</STRONG><BR>
              以上信息我网未做亲自核实，我网将力求但不能确保信息的准确性。您若发现此为虚假信息立即投诉，一经核实将删除其会员帐号！<BR>
              声明：此条信息不得转载到其他网站，违者将被子停止帐号服务。 
              <BR>
            </P></TD></TR></TBODY></TABLE></TD>
  </TR></TBODY></TABLE>
</body>
</HTML>
