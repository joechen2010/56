<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/function.asp"-->
<%
Dim GsID
GsID=SafeRequest("GsID",1)
if GsID="" then
     call WriteErrMsg2()
end if
set rs_q=server.createobject("adodb.recordset")
sql="select * from qyml where id="&GsID&""
rs_q.open sql,conn,1,1
if rs_q.eof and rs_q.bof then
     call WriteErrMsg2()
else
%>
<HTML>
<HEAD>
<TITLE><%=rs_q("qymc")%> - ����ͨ������ҵ - ���ŵ���</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
<style>
a{
	font-size: 10pt;
	color: #000000;
	text-decoration: none;
	FONT-FAMILY: "����"
}
a:hover {
	color: #FF0000;
	text-decoration: none;
}
td {
	font-size: 10pt;
	text-decoration: blink;
}
</style>
</HEAD>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="head.htm" -->
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="169" background="images/mb_01.jpg"><DIV class=tile> ����<font color="#FF0000" size="+3"><STRONG><%=rs_q("qymc")%></STRONG></font>
            </DIV></td>
  </tr>
</table>
<table width="1003" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="200" align="center" valign="top">
<!--#include file="left.htm" -->
      <table width="160" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="20">&nbsp;</td>
          <td width="120">&nbsp;</td>
          <td width="20">&nbsp;</td>
        </tr>
        <tr> 
          <td background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td background="images/bk_top2.jpg"> <DIV class=tile2 
              align=center><STRONG> ��ϵ����</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr> 
          <td colspan="3" valign="top" nowrap class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999; border-top: 1px solid #999999;" width="160">&nbsp;</td>
        </tr>
        <tr> 
          <td  width="160" height="25" colspan="3" nowrap bgcolor="#F6F6F6" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">��ַ��<%=rs_q("address")%></td>
        </tr>
        <tr> 
          <td  width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;" td>�ʱࣺ<font face="Arial"><%=rs_q("post")%></font></td>
        </tr>
        <tr> 
          <td  width="160" height="25" colspan="3" bgcolor="#F6F6F6"  class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">�绰��<font face="Arial"> 
            <% Response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3")%>
            </font></td>
        </tr>
        <tr> 
          <td  width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;">���棺<font face="Arial"> 
            <% Response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
            </font></td>
        </tr>
        <tr> 
          <td width="160" height="25" colspan="3" bgcolor="#F6F6F6" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;"><font face="Arial">E-mail��</font><font face="Arial"><a href="mailto:<%=rs_q("email")%>"><%=rs_q("email")%></a></font></td>
        </tr>
        <tr> 
          <td width="160" height="25" colspan="3" class=S2 style="border-left: 1px solid #999999;border-right:1px solid #999999;border-bottom: 1px solid #999999">��ҳ��<font face="Arial"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a></font></td>
        </tr>
      </table>
    </td>
    <td width="803" align="center" valign="top"> <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="images/bk_top2.jpg"><img src="images/bk_top1.jpg" width="20" height="31"></td>
          <td width="731" background="images/bk_top2.jpg"> <DIV class=tile2 
              ><STRONG> ���ŵ���</STRONG></DIV></td>
          <td align="right" background="images/bk_top2.jpg"><img src="images/bk_top3.jpg" width="20" height="31"></td>
        </tr>
        <tr align="center"> 
          <td colspan="3"  class=C vAlign=top borderColor=#cccccc style="border: 1px solid #999999;"><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>
            <table class=style height=30 cellspacing=0 cellpadding=0 width="100%" 
            align=center bgcolor=#ffffff background=images/bj.jpg 
            border=0>
                      <tbody> 
              <tr> 
                          <td bgcolor=#BDEEDE><span class=zi1>����Ա���¼ ��Ա����� 1 
                            ��</span></td>
              </tr>
              </tbody>
            </table>
                    <table class=style cellspacing=1 cellpadding=6 width="100%" 
            align=center border=0>
                      <tr> 
                        <td width="24%" bgcolor=#E4F8F1 align="left">���ڳ�������������ķ���</td>
                        <td width="61%"> 
                          <%
				   if rs_q("flag")=0 then
	  ss="��ѻ�Ա"
   elseif rs_q("flag")=1 then
	  ss="ͭ�ƻ�Ա"
   elseif rs_q("flag")=2 then
	  ss="���ƻ�Ա"
   elseif rs_q("flag")=3 then
	  ss="���ƻ�Ա"
   end if
   response.write ss
				%>
                  <br>
                  (<%=rs_q("BeginDate")%> ~ <%=rs_q("enddate")%>)</td>
              </tr>
              <tbody> 
                        <tr> 
                          <td width="24%" bgcolor=#E4F8F1 align="left"> ע���Ա��ʱ��</td>
                          <td width="61%"><%=rs_q("idate")%></td>
              </tr>
                        <tr> 
                          <td bgcolor=#E4F8F1> �����¼��ʱ��</td>
                          <td><%=rs_q("LastLoginTime")%></td>
              </tr>
                        <tr> 
                          <td bgcolor=#E4F8F1 align="left"> �ѷ�����������Ϣ����</td>
                          <td>�ѷ�����������Ϣ������<br>
                  �����̻���<font color=#ff0000>
                  <%set rs_p=conn.execute("select count(*) from spzs where gsid='"&rs_q("user")&"'")
				  response.write rs_p(0)
				 ' rs_p.close
				 'set rs_p=nothing%>
                  </font><br>
                  չʾ��Ʒ��<font color=#ff0000><%=conn.execute("select count(*) from info where gsid='"&rs_q("user")&"'")(0)%></font></td>
              </tr>
              </tbody> 
            </table>
                    <table class=style height=30 cellspacing=0 cellpadding=0 width="100%" 
            align=center bgcolor=#fafafa border=0>
                      <tbody> 
              <tr> 
                          <td bgcolor=#BCEDDC><span 
            class=zi1>��֤�鼰����</span></td>
              </tr>
              </tbody>
            </table>
                    <table class=style cellspacing=2 cellpadding=2 width="100%" 
            align=center border=0>
                      <tbody> 
              <tr bgcolor=#f5f5f5> 
                          <td width="49%" height="89" align=middle> 
                            <%set QyzsRs=Conn.Execute("Select * from qyzs where gsid='"&rs_q("user")&"'")
				   if QyzsRs.eof and QyzsRs.bof then
				       response.write "�ù�˾��δ������ҵ֤��!"
				   else
				   %><table class=style cellspacing=3 cellpadding=3 width="100%" 
                  border=0>
                    <tbody>
					<tr>  
					<%
					dim i
					i=0
					do while not QyzsRs.eof%>
                    
                      <td width='50%'><img src='../../Member/<%=QyzsRs("zsurl")%>'></td>
					  <%
					  i=i+1
					  if i mod 2=0 then Response.Write "</tr><tr>"
					  QyzsRs.movenext
					  loop
					  %>
                    </tr>
                    </tbody>
                  </table>
                <%
				   end if
				   QyzsRs.close
				   set QyzsRs=nothing
				  %>
				  </td>
              </tr>
              </tbody>
            </table>
                    <table class=style cellspacing=3 cellpadding=3 width="100%" 
            align=center border=0>
                      <tbody> 
              <tr> 
                          <td 
                  style="FONT-SIZE: 10pt; LINE-HEIGHT: 23px" bgcolor=#D9F4EB>��ע�⣺<br>
                            ��������������ʾ������Ϣ���������������������յ������������֤��ʹ��������Ϣʱ�������֮�����Ҳ����������Ϣ��׼ȷ�ԣ������Ժͼ�ʱ�ԣ�Ҳ���е�����ߵ��κ���ҵ���ա���ˣ��������������е�����˶���������ʹ���κ��𺦡���������������ȫ���򲿷�ɾ�����������Ȩ��������ҳ�����ݽ�ֹ�Ը�����ʽ��ȫ���򲿷ָ���ʹ�á� 
                          </td>
              </tr>
              </tbody>
            </table>
            <P>&nbsp;</P></TD></TR></TBODY></TABLE></td>
        </tr>
        <tr> 
          <td height="30" colspan="3">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="Copyright.htm" -->
<%end if%>
</body>
</html>