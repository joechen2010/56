
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
          <TD class=title02 bgColor=#ffeee0 colSpan=4>���˵�װ</TD>
        </TR>
        <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from banyun_info a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		  if not (rs.eof and rs.bof) then
		%>
        <TR> 
          <TD width="16%" bgColor=#f7f7f7>���⣺</TD>
          <TD colspan="3"><%=rs("title")%></TD>
        </TR>
        <TR> 
          <TD bgColor=#e7e7e7 colSpan=4 height=1></TD>
        </TR>
        <TR> 
          <TD bgColor=#f7f7f7 width="16%">����ʡ�У�</TD>
          <TD colspan="3"><%=rs("a.province")%> <%=rs("a.city")%> <%=rs("a.area")%></TD>
        </TR>
        <TR> 
          <TD bgColor=#e7e7e7 colSpan=4 height=1></TD>
        </TR>
        <TR> 
          <TD bgColor=#f7f7f7 width="16%">��Ч�ڣ�</TD>
          <TD width="22%"><%=rs("validate")%>��</TD>
          <TD bgColor=#f7f7f7 width="12%">����ʱ�䣺</TD>
          <TD width="50%"><%=rs("addtime")%></TD>
        </TR>
        <TR> 
          <TD bgColor=#f7f7f7 width="16%">��ϵ�ˣ�</TD>
          <TD width="22%"><%=rs("name")%></TD>
          <TD bgColor=#f7f7f7 width="12%">�绰��</TD>
          <TD width="50%"><%=rs("phone2")%>-<%=rs("phone3")%></TD>
        </TR>
        <TR> 
          <TD bgColor=#f7f7f7 width="16%">���ݣ�</TD>
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
            <P><STRONG>168������Ϣ�����ڱ�����Ϣȷ�Ե�������</STRONG><BR>
              ������Ϣ����δ�����Ժ�ʵ�����������󵫲���ȷ����Ϣ��׼ȷ�ԡ��������ִ�Ϊ�����Ϣ����Ͷ�ߣ�һ����ʵ��ɾ�����Ա�ʺţ�<BR>
              ������������Ϣ����ת�ص�������վ��Υ�߽�����ֹͣ�ʺŷ��� 
              <BR>
            </P></TD></TR></TBODY></TABLE></TD>
  </TR></TBODY></TABLE>
</body>
</HTML>
