<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%infoid=request("infoid")
  user=request("user")
%>
<HTML>
<HEAD>
<title>��Ա��ϸ��Ϣ</title>
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
<script language="javascript">
function openwindow(url,width,height) { 	
left1 = (screen.width-800)/2; 	
top1 = (screen.height-650)/2; 	
window.open(url,"","width=" + width + ",status=no,scrollbars=yes,height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); 
}

</script>
</HEAD>
<body leftmargin="0" topmargin="0">
 <TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>
        <TR align="center">
          <TD width="158" height="30"><a href="site.asp?infoid=<%=request("infoid")%>&user=<%=request("user")%>">��������</a></TD>
          <TD width="136"><a href="site.asp?infoid=<%=request("infoid")%>&user=<%=request("user")%>&action=goods">������Ϣ</a></TD>
          <TD width="139"><a href="site.asp?infoid=<%=request("infoid")%>&user=<%=request("user")%>&action=zx">����ר��</a></TD>
          <TD width="127"><a href="site.asp?infoid=<%=request("infoid")%>&user=<%=request("user")%>&action=file">��������</a></TD>
          <TD width="123" height="30">
		  <a href="site.asp?infoid=<%=request("infoid")%>&user=<%=request("user")%>&action=ly">��������</a></TD>
          <TD width="95"><a href="../index.asp">������ҳ</a></TD>
        </TR>
</TABLE>
<%if session("user")<>"" then%>
<%if request("action")="" then%>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0>
  <TR>
    <TD align="center" vAlign=top>
      <TABLE width=96% border=0 align=center cellPadding=5 cellSpacing=0>
        <TR>
          <TD bgColor=#ff811e colSpan=4 height=2></TD></TR>
        <TR>
          <TD class=title02 bgColor=#ffeee0 colSpan=4>&nbsp;</TD>
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
          <TD colspan="3"><%=rs("qymc")%></TD>
        </TR>
        <TR>
          <TD bgColor=#e7e7e7 colSpan=4 height=1></TD></TR>
        <TR>
          <TD bgColor=#f7f7f7>���ڳ��У�</TD>
          <TD colspan="3"><%=rs("province")%>&nbsp;<%=rs("city")%>&nbsp;<%=rs("area")%></TD>
        </TR>
        <TR>
          <TD bgColor=#e7e7e7 colSpan=4 height=1></TD></TR>
        <TR>
          <TD bgColor=#f7f7f7>��˾��ַ��</TD>
          <TD colspan="3"><%=rs("address")%></TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>�ʱࣺ</TD>
          <TD colspan="3"><%=rs("post")%></TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>��˾��飺</TD>
          <TD colspan="3"><%=rs("qyjj")%></TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>��ϵ�ˣ�</TD>
          <TD width="14%" align="center"><%=rs("name")%></TD>
          <TD width="12%" bgColor=#f7f7f7>E-mail:</TD>
          <TD><%=rs("email")%></TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>�绰��</TD>
          <TD align="center"><%=rs("phone1")%></TD>
          <TD bgColor=#f7f7f7>���棺</TD>
          <TD><%=rs("fax1")%></TD>
        </TR>
        <TR>
          <TD bgColor=#f7f7f7>ע��ʱ�䣺</TD>
          <TD align="center"><%=rs("idate")%></TD>
          <TD bgColor=#f7f7f7>��ַ��</TD>
          <TD><a href="<%=rs("web")%>" target="_blank"><%=rs("web")%></a></TD>
        </TR>
        <TR>
          <TD bgColor=#e7e7e7 colSpan=4 height=1></TD></TR>
        <TR>
          <TD bgColor=#e7e7e7 colSpan=4 height=1></TD></TR>
        <TR>
        </TR>
        <TR><TD bgColor=#ff811e colSpan=4 height=2></TD></TR>
   </TABLE>
</TD>
  </TR>
</TABLE>
<%end if%>
<%if request("action")="goods" then%>
<TABLE cellSpacing=0 cellPadding=5 width=778 align=center border=0>
        <TR>
          <TD><!--#include file="site1.asp"--></TD>
</TR>
</TABLE>
<%end if%>
<%if request("action")="zx" then%>
<TABLE cellSpacing=0 cellPadding=5 width=778 align=center border=0>
   <TR>
    <TD><!--#include file="site2.asp"--></TD>
  </TR>
</TABLE>		  
<%end if%>
<%if request("action")="file" then%>
<TABLE cellSpacing=0 cellPadding=5 width=778 align=center border=0>
   <TR>
    <TD><!--#include file="site3.asp"--></TD>
  </TR>
</TABLE>		  
<%end if%>
<%if request("action")="ly" then%>
<TABLE cellSpacing=0 cellPadding=5 width=778 align=center border=0>
   <TR>
    <TD><!--#include file="site4.asp"--></TD>
  </TR>
</TABLE>		  
<%end if%>
<%end if%>
   <%if session("user")="" then%>
<TABLE cellSpacing=0 cellPadding=5 width=778 align=center border=0>
        <TR>
          <TD>   
   <% Response.Write("<font color='#999999'>Ϊ��֤��Ϣ�����ߵ���˽Ȩ������Ҫ<a href='../Login/login.asp' target='_blank'>��¼</a>�ſ��Բ鿴��ϵ��ʽ</font>")
    Response.Write("<br><a href='../Reg/User_Reg.asp' target='_blank'>δע���û�����>></a>")
	%>
		  </TD>
		</TR>
</TABLE>	
	<%end if%>
<TABLE cellSpacing=0 cellPadding=5 width=778 align=center border=0>
        <TR>
          <TD>
            <P><STRONG>168������Ϣ�����ڱ�����Ϣȷ�Ե�������</STRONG><BR>
              ������Ϣ����δ�����Ժ�ʵ�����������󵫲���ȷ����Ϣ��׼ȷ�ԡ��������ִ�Ϊ�����Ϣ����Ͷ�ߣ�һ����ʵ��ɾ�����Ա�ʺţ�<BR>
              ������������Ϣ����ת�ص�������վ��Υ�߽�����ֹͣ�ʺŷ��� 
              <BR>
            </P>
		  </TD>
		</TR>
</TABLE>
<!--#include file="../inc/down.htm"-->
</body>
</HTML>
