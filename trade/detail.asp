<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<HTML>
<HEAD>
<title>������Ϣ</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/page.css" type=text/css rel=stylesheet>
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
<table width=480 border=0 align=center cellpadding=2 cellspacing=2>
  <tbody> 
  <tr> 
    <td colspan=6 height=2><embed src="../truck/images/banner.swf" width="480" height="60" type="application/x-shockwave-flash">
      </embed></td>
  </tr>
  <tr align="center"> 
    <td class=title02 height="20" colspan=6 bgcolor=#B1EBC6><b>������Ϣ</b></td>
  </tr>
  <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from gq_info a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
  <tr> 
    <td width="16%"><b>���⣺</b></td>
    <td colspan="5"><%=rs("title")%></td>
  </tr>
  <tr> 
    <td><b>��Ϣ���ͣ�</b></td>
    <td colspan="5"><%=rs("infotype")%></td>
  </tr>
  <tr> 
    <td><b>����ʡ�У�</b></td>
    <td colspan="5"><%=rs("a.province")%>&nbsp;&nbsp;<%=rs("a.city")%>&nbsp;&nbsp;<%=rs("a.area")%></td>
  </tr>
  <tr> 
    <td><b>��Ч�ڣ�</b></td>
    <td colspan="5"><%=rs("validate")%>��</td>
  </tr>
  <tr> 
    <td><b>����ʱ�䣺</b></td>
    <td colspan="5"><%=rs("addtime")%></td>
  </tr>
  <tr> 
    <td><b>��ϵ�ˣ�</b></td>
    <td colspan="5"><%=rs("name")%></td>
  </tr>
  <tr> 
    <td><b>�绰��</b></td>
    <td colspan="5"><%=rs("phone2")%>-<%=rs("phone3")%></td>
  </tr>
  <tr> 
    <td><b>���ݣ�</b></td>
    <td colspan="5"><%=rs("content")%></td>
  </tr>
  <tr align="center"> 
    <td colspan=6 bgcolor="#e7e7e7" height="20">168������Ϣ�����ڱ�����Ϣȷ�Ե�������</td>
  </tr>
  <tr> 
    <td colspan=6><font color="#999999">������Ϣ����δ�����Ժ�ʵ�����������󵫲���ȷ����Ϣ��׼ȷ�ԡ��������ִ�Ϊ�����Ϣ����Ͷ�ߣ�һ����ʵ��ɾ�����Ա�ʺţ�<br>
      ������������Ϣ����ת�ص�������վ��Υ�߽�����ֹͣ�ʺŷ��� </font></td>
  </tr>
  <tr> 
    <td colspan=6 height="20" bgcolor="#B1EBC6">[<a id="HyperLink1" target="_blank">�鿴�û�Ա��վ</a>] 
      [<a href="JavaScript:window.print()">��ӡ��ҳ</a>] [<a href="javascript:window.close()">�رմ���</a>]</td>
  </tr>
  </tbody> 
</table>
</body>
</HTML>
