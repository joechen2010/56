<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
        <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from file_info a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
<HTML>
<HEAD>
<title>��������--<%=rs("carnum")%></title>
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
<table width=480 border=0 align=center cellpadding=0 cellspacing=0>
  <tbody> 
  <tr> 
    <td height=20><embed src="../flash/banner.swf" width="480" height="60" type="application/x-shockwave-flash">
      </embed></td>
  </tr>
  <tr bgcolor="#B1EBC6" align="center"> 
    <td class=title02 height="20"><b>��������--<%=rs("carnum")%></b></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7> 
      <table border=0 align=center cellpadding=2 cellspacing=2 width=100%>
        <tbody> 
        <tr> 
          <td width="16%" bgcolor=#f7f7f7><b>���ƺţ�</b></td>
          <td width="33%" align="center"><%=rs("carnum")%></td>
          <td width="16%" bgcolor=#f7f7f7><b>������</b></td>
          <td colspan="3" width="35%"><%=rs("qymc")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7 width="16%"><b>Ʒ�ƣ�</b></td>
          <td align="center" width="33%"><%=rs("carpp")%></td>
          <td bgcolor=#f7f7f7 width="16%"><b>������</b></td>
          <td colspan="3" width="35%"><%=rs("city")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7 width="16%"><b>���ͣ�</b></td>
          <td align="center" width="33%"><%=rs("cartype")%></td>
          <td bgcolor=#f7f7f7 width="16%"><b>���أ�</b></td>
          <td colspan="3" width="35%"><%=rs("carload")%>��</td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7 width="16%"><b>������</b></td>
          <td align="center" width="33%"><%=rs("carlong")%>��</td>
          <td bgcolor=#f7f7f7 width="16%"><b>����ʱ�䣺</b></td>
          <td colspan="3" width="35%"><%=rs("time")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7 width="16%"><b>��Ҫ˾����</b></td>
          <td align="center" width="33%"><%=rs("b.siji")%></td>
          <td bgcolor=#f7f7f7 width="16%"><b>�ó�״����</b></td>
          <td colspan="3" width="35%"><%=rs("content")%></td>
        </tr>
        <tr> 
          <td bgcolor=#f7f7f7 width="16%"><b>����ʱ�䣺</b></td>
          <td colspan="5"><%=rs("addtime")%></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  <tr align="center"> 
    <td bgcolor=#EFEFEF height=27>168������Ϣ�����ڱ�����Ϣȷ�Ե�������</td>
  </tr>
  <tr> 
    <td height=54><font color="#999999">������Ϣ����δ�����Ժ�ʵ�����������󵫲���ȷ����Ϣ��׼ȷ�ԡ��������ִ�Ϊ�����Ϣ����Ͷ�ߣ�һ����ʵ��ɾ�����Ա�ʺţ�<br>
      ������������Ϣ����ת�ص�������վ��Υ�߽�����ֹͣ�ʺŷ���</font></td>
  </tr>
  <tr align="center" bgcolor="#B1EBC6"> 
    <td height=20>[<a id="HyperLink1" target="_blank">�鿴�û�Ա��վ</a>] [<a href="JavaScript:window.print()">��ӡ��ҳ</a>] 
      [<a href="javascript:window.close()">�رմ���</a>]</td>
  </tr>
  </tbody> 
</table>
</body>
</HTML>
