
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
<table width=500 border=0 align=center cellpadding=2 cellspacing=2>
  <tbody> 
  <tr> 
    <td colspan=2 height=2><embed title="�л�������" src="images/banner.swf" width="500" height="60" type="application/x-shockwave-flash">
      </embed></td>
  </tr>
  <tr> 
    <td class=title02 colspan=2 height="20" bgcolor="#B1EBC6" align="center"><b>������Ϣ</b></td>
  </tr>
  <%
        dim infoID
         infoID=Request.QueryString("infoID")
        if isnumeric(infoID)=0 or infoID="" then
         response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from info2 a,qyml b where a.infoid="&request("infoid")&" and a.gsid=b.user"
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		  if not (rs.eof and rs.bof) then
		%>
  <tr> 
    <td width="16%" bgcolor=#f7f7f7><b>�����ص㣺</b></td>
    <td><%=rs("a.province")%>&nbsp;&nbsp;<%=rs("a.city")%>&nbsp;&nbsp;<%=rs("a.area")%></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>����ص㣺</b></td>
    <td><%=rs("province2")%>&nbsp;&nbsp;<%=rs("city2")%>&nbsp;&nbsp;<%=rs("area2")%></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>�������ͣ�</b></td>
    <td><%=rs("cartype")%> <b>���أ�</b><%=rs("carload")%>�� <b>������</b><%=rs("carlong")%>��</td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>��Ч�ڣ�</b></td>
    <td><%=rs("validate")%>��</td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>��ע��</b></td>
    <td><%=rs("content")%></td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>��ϵ�ˣ�</b></td>
    <td> 
      <%if session("user")<>"" then%>
      <%=rs("qymc")%>&nbsp;&nbsp;<%=rs("name")%> 
      <%else
		    response.Write "<font color='#FF0000'>����û�е�½����ϵ�˲���ʾ</font>"
			end if
		  %>    </td>
  </tr>
  <tr> 
    <td bgcolor=#f7f7f7><b>�绰��</b></td>
    <td> 
      <%if session("user")<>"" then%>
      <%=rs("phone2")%>-<%=rs("phone3")%> 
      <%else
		    response.Write "<font color='#FF0000'>ֻ�л�Ա��¼����ܲ鿴</font>"
			end if
		  %>
      <%if session("user")="" then%>
      <% Response.Write(" <a href='../Reg/User_Reg.asp' target='_blank'><font color='#0000FF'>�����ǻ�Ա��ע��</font>")
    Response.Write(" <a href='../Login/login.asp' target='_blank'><font color='#0000FF'>�ǻ�Ա���¼</font></a>")
	%>
      <%end if%>
    &nbsp;&nbsp;<%if session("user")<>"" then%>
      <%=rs("mobile")%> 
      <%else
		    response.Write "����û�е�½���绰����ʾ"
			end if
		  %></td>
  </tr>
  
  <tr> 
    <td bgcolor=#f7f7f7><b>����ʱ�䣺</b></td>
    <td><%=rs("addtime")%></td>
  </tr>
  <%end if%>
  <tr> 
    <td bgcolor=#e7e7e7 colspan=2 height=1></td>
  </tr>
  <tr> 
    <td colspan=2 height=20 align="center" bgcolor="#EFEFEF">168������Ϣ�����ڱ�����Ϣȷ�Ե�������</td>
  </tr>
  <tr> 
    <td colspan=2><font color="#999999">������Ϣ����δ�����Ժ�ʵ�����������󵫲���ȷ����Ϣ��׼ȷ�ԡ��������ִ�Ϊ�����Ϣ����Ͷ�ߣ�һ����ʵ��ɾ�����Ա�ʺţ�<br>
      ������������Ϣ����ת�ص�������վ��Υ�߽�����ֹͣ�ʺŷ��� </font></td>
  </tr>
  <tr bgcolor="#B1EBC6" align="center"> 
    <td colspan=2 height="20">[<a id="HyperLink1" target="_blank">�鿴�û�Ա��վ</a>] 
      [<a href="JavaScript:window.print()">��ӡ��ҳ</a>] [<a href="javascript:window.close()">�رմ���</a>]</td>
  </tr>
  </tbody> 
</table>
</body>
</HTML>
