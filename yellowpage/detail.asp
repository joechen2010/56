
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<HTML>
<HEAD>
<title>��Ա��ϸ��Ϣ</title>
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
          <TD class=title02 width=289>����Ա��Ϣ</TD>
          <TD align=right width=200></TD></TR></TBODY></TABLE>
      <TABLE width=96% border=0 align=center cellPadding=5 cellSpacing=0>
        <TBODY>
        <TR>
          <TD bgColor=#ff811e colSpan=6 height=2></TD></TR>
        <TR>
          <TD class=title02 bgColor=#ffeee0 colSpan=6>�鿴�û�Ա��վ </TD>
        </TR>
		<%
        dim infoid
         infoid=Request.QueryString("infoid")
        if isnumeric(infoid)=0 or infoid="" then
         response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from qyml where id="&request("infoid")&""
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
        <TR>
          <TD width="16%" bgColor=#f7f7f7>��ҵ���ƣ�</TD>
          <TD width="14%" align="center"><%=rs("qymc")%></TD>
          <TD width="12%" align="center" bgColor=#f7f7f7>&nbsp;</TD>
          <TD colspan="2" align="center">&nbsp;</TD>
          <TD width="39%">&nbsp;</TD>
        </TR>
        <TR>
          <TD bgColor=#e7e7e7 colSpan=6 height=1></TD></TR>
        <TR>
          <TD bgColor=#f7f7f7>���ڳ��У�</TD>
          <TD align="center"><%=rs("province")%></TD>
          <TD bgColor=#f7f7f7 align="center"><%=rs("city")%></TD>
          <TD width="11%" align="center"><%=rs("area")%></TD>
          <TD width="8%">&nbsp;</TD>
          <TD>&nbsp;</TD>
        </TR>
        <TR>
          <TD bgColor=#e7e7e7 colSpan=6 height=1></TD></TR>
        <TR>
          <TD bgColor=#f7f7f7>��˾��ַ��</TD>
          <TD colspan="4"><%=rs("address")%></TD>
          <TD align="center"></TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>�ʱࣺ</TD>
          <TD align="center"><%=rs("post")%></TD>
          <TD bgColor=#f7f7f7>&nbsp;</TD>
          <TD>&nbsp;</TD>
          <TD>&nbsp;</TD>
          <TD>&nbsp;</TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>��˾��飺</TD>
          <TD colspan="5"><%=rs("qyjj")%></TD>
          </TR>
        <TR>
          <TD bgColor=#f7f7f7>��ϵ�ˣ�</TD>
          <TD align="center"><%=rs("name")%></TD>
          <TD bgColor=#f7f7f7>E-mail:</TD>
          <TD><%=rs("email")%></TD>
          <TD bgColor=#f7f7f7>��ַ��</TD>
          <TD><a href="<%=rs("web")%>" target="_blank"><%=rs("web")%></a></TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>�绰��</TD>
          <TD align="center"><%=rs("phone1")%></TD>
          <TD bgColor=#f7f7f7>���棺</TD>
          <TD colspan="2"><%=rs("fax1")%></TD>
          <TD>&nbsp;</TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>ע��ʱ�䣺</TD>
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
            <P><STRONG>168������Ϣ�����ڱ�����Ϣȷ�Ե�������</STRONG><BR>
              ������Ϣ����δ�����Ժ�ʵ�����������󵫲���ȷ����Ϣ��׼ȷ�ԡ��������ִ�Ϊ�����Ϣ����Ͷ�ߣ�һ����ʵ��ɾ�����Ա�ʺţ�<BR>
              ������������Ϣ����ת�ص�������վ��Υ�߽�����ֹͣ�ʺŷ��� 
              <BR>
            </P></TD></TR></TBODY></TABLE></TD>
  </TR></TBODY></TABLE>
<!--#include file="../inc/down.htm"-->
</body>
</HTML>
