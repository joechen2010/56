
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<HTML>
<HEAD>
<title>会员详细信息</title>
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
<!--#include file="../inc/top.htm"-->
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="center"><iframe id="ifrsearch" src="../search/index_car.asp" width="100%" frameborder="0" scrolling="no" height="90"></iframe></td>
  </tr>
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
      <TABLE cellSpacing=0 cellPadding=0 width=96% align=center border=0>
        <TBODY>
        <TR>
          <TD colSpan=2>&nbsp;</TD>
        </TR>
        <TR>
          <TD colSpan=2 height=8></TD></TR>
        <TR>
          <TD class=title02 width=289>　会员信息</TD>
          <TD align=right width=200></TD></TR></TBODY></TABLE>
      <TABLE width=96% border=0 align=center cellPadding=5 cellSpacing=0>
        <TBODY>
        <TR>
          <TD bgColor=#ff811e colSpan=6 height=2></TD></TR>
        <TR>
          <TD class=title02 bgColor=#ffeee0 colSpan=6>查看该会员网站 </TD>
        </TR>
		<%
        dim infoid
         infoid=Request.QueryString("infoid")
        if isnumeric(infoid)=0 or infoid="" then
         response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from qyml where id="&request("infoid")&""
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
        <TR>
          <TD width="16%" bgColor=#f7f7f7>企业名称：</TD>
          <TD width="14%" align="center"><%=rs("qymc")%></TD>
          <TD width="12%" align="center" bgColor=#f7f7f7>&nbsp;</TD>
          <TD colspan="2" align="center">&nbsp;</TD>
          <TD width="39%">&nbsp;</TD>
        </TR>
        <TR>
          <TD bgColor=#e7e7e7 colSpan=6 height=1></TD></TR>
        <TR>
          <TD bgColor=#f7f7f7>所在城市：</TD>
          <TD align="center"><%=rs("province")%></TD>
          <TD bgColor=#f7f7f7 align="center"><%=rs("city")%></TD>
          <TD width="11%" align="center"><%=rs("area")%></TD>
          <TD width="8%">&nbsp;</TD>
          <TD>&nbsp;</TD>
        </TR>
        <TR>
          <TD bgColor=#e7e7e7 colSpan=6 height=1></TD></TR>
        <TR>
          <TD bgColor=#f7f7f7>公司地址：</TD>
          <TD colspan="4"><%=rs("address")%></TD>
          <TD align="center"></TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>邮编：</TD>
          <TD align="center"><%=rs("post")%></TD>
          <TD bgColor=#f7f7f7>&nbsp;</TD>
          <TD>&nbsp;</TD>
          <TD>&nbsp;</TD>
          <TD>&nbsp;</TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>公司简介：</TD>
          <TD colspan="5"><%=rs("qyjj")%></TD>
          </TR>
        <TR>
          <TD bgColor=#f7f7f7>联系人：</TD>
          <TD align="center"><%=rs("name")%></TD>
          <TD bgColor=#f7f7f7>E-mail:</TD>
          <TD><%=rs("email")%></TD>
          <TD bgColor=#f7f7f7>网址：</TD>
          <TD><a href="<%=rs("web")%>" target="_blank"><%=rs("web")%></a></TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>电话：</TD>
          <TD align="center"><%=rs("phone1")%></TD>
          <TD bgColor=#f7f7f7>传真：</TD>
          <TD colspan="2"><%=rs("fax1")%></TD>
          <TD>&nbsp;</TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>注册时间：</TD>
          <TD align="center"><%=rs("idate")%></TD>
          <TD bgColor=#f7f7f7>&nbsp;</TD>
          <TD>&nbsp;</TD>
          <TD>&nbsp;</TD>
          <TD>&nbsp;</TD>
        </TR>
        <TR>
          <TD bgColor=#e7e7e7 colSpan=6 height=1></TD></TR>
        
        <TR>
          <TD bgColor=#e7e7e7 colSpan=6 height=1></TD></TR>
        <TR>
          <TD bgColor=#f7f7f7>&nbsp;</TD>
          <TD>&nbsp;</TD>
        </TR>
        <TR>
          <TD bgColor=#ff811e colSpan=6 height=2></TD></TR></TBODY></TABLE>
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
<!--#include file="../inc/down.htm"-->
</body>
</HTML>
