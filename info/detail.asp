
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
<table width=96% border=0 align=center cellpadding=5 cellspacing=5>
  <tbody> 
  <tr> 
    <td bgcolor=#ff811e colspan=6 height=2></td>
  </tr>
  <tr align="center"> 
    <td class=title02 bgcolor=#ffeee0 colspan=6><b>�ⷿ��Ϣ</b></td>
  </tr>
  <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from kf_info a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
  <tr> 
    <td width="16%" bgcolor=#f7f7f7><b>��Ϣ���ͣ�</b></td>
    <td width="9%" align="center"><%=rs("infotype")%></td>
    <td width="10%" align="center" bgcolor=#f7f7f7><b>��ϵ�ˣ�</b></td>
    <td colspan="3" align="center"><%=rs("name")%></td>
  </tr>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=6 height=1><b></b></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>�ⷿ���ͣ�</b></td>
    <td align="center"><%=rs("kftype")%></td>
    <td bgcolor=#f7f7f7 align="center"><b>�绰��</b></td>
    <td align="center" colspan="3"><%=rs("phone2")%>-<%=rs("phone3")%></td>
  </tr>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=6 height=1><b></b></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>�ֿ�ص㣺</b></td>
    <td colspan="5"><%=rs("a.province")%><%=rs("a.city")%><%=rs("a.area")%></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>�ִ�������</b></td>
    <td colspan="5"><%=rs("tj1")%>&nbsp;<%=rs("tj2")%>&nbsp;<%=rs("tj3")%></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>�����</b></td>
    <td align="center"><%=rs("mji")%>ƽ����</td>
    <td bgcolor=#f7f7f7 align="center"><b>�۸�</b></td>
    <td colspan="3"><%=rs("prices")%>Ԫ/��</td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>��Ч�ڣ�</b></td>
    <td align="center"><%=rs("validate")%>��</td>
    <td bgcolor=#f7f7f7 align="center"><b>��ע��</b></td>
    <td colspan="3"><%=rs("content")%></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>����ʱ�䣺</b></td>
    <td colspan="5"><%=rs("addtime")%></td>
  </tr>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=6 height=1></td>
  </tr>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=6 height=1></td>
  </tr>
  <tr> 
    <td bgcolor=#ff811e colspan=6 height=2></td>
  </tr>
  </tbody> 
</table>
</body>
</HTML>
