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
<LINK href="images/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<table cellspacing=0 cellpadding=0 width="100%" border=0>
  <tbody> 
  <tr> 
    <td valign=top width=189 bgcolor=#fcf6e8> 
      <table height=69 cellspacing=0 cellpadding=0 width="100%" 
      background=images/img_bg_9.gif border=0>
        <tbody> 
        <tr> 
          <td valign="bottom"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="20"> 
                  <div align="center"><font color=#ff0000>����ͨ��Ա��������<%=rs_q("idate")%></font></div>
                </td>
              </tr>
            </table>
          </td>
          <td align=right width=10><img height=69 src="images/img_11.gif" 
            width=10></td>
        </tr>
        </tbody> 
      </table>
      <!--#include file="left.htm"-->
      <table cellspacing=0 cellpadding=0 width=100% border=0>
        <tbody> 
        <tr> 
          <td align=middle height="15">&nbsp;</td>
        </tr>
        <tr> 
          <td align=middle> 
            <table width="96%" border="0" cellpadding="0" cellspacing="0" class=S2 style="border-top:1px dotted #FF6600" align="center">
              <tr> 
                <td width="22%" nowrap class=S2 height="28">��ַ��</td>
                <td width="78%" colspan="3" class=S2 height="28"><%=rs_q("address")%></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">�ʱࣺ</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"><%=rs_q("post")%></font></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">�绰��</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"> 
                  <% Response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3")%>
                  </font></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">���棺</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"> 
                  <% Response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
                  </font></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">��ҳ��</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a>&nbsp; 
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td valign=top align=middle> 
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
        <tbody> 
        <tr> 
          <td valign=top height=12><img height=10 src="images/img_12.gif" 
            width=11></td>
          <td valign=top align=right><img height=10 src="images/img_14.gif" 
            width=10></td>
        </tr>
        </tbody> 
      </table>
      <table cellspacing=0 cellpadding=0 width="96%" 
      background=images/img_bg_31.gif border=0>
        <tbody> 
        <tr> 
          <td width="4%" height=29><img height=29 src="images/img_30.gif" 
            width=13></td>
          <td class=font14 width="92%"><strong>���ŵ���</strong></td>
          <td align=right width="4%"><img height=29 src="images/img_33.gif" 
            width=10></td>
        </tr>
        </tbody> 
      </table>
      <table class=style height=40 cellspacing=0 cellpadding=0 width="96%" 
            align=center 
            border=0>
        <tr>
          <td height="20"></td>
        </tr>
        <tbody> 
        <tr> 
          <td bgcolor=#FFCC00 height="30"><span class=zi1>����Ա���¼ ��Ա����� 1 ��</span></td>
        </tr>
        </tbody> 
      </table>
      <table class=style cellspacing=1 cellpadding=6 width="96%" 
            align=center border=0>
        <tr bgcolor=#f5f5f5> 
          <td width="24%" bgcolor=#fff4e5 align="left">���ڳ�������������ķ���</td>
          <td width="61%" bgcolor=#F7F7F7> <font color="#FF0000"> 
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
            </font> (<%=rs_q("BeginDate")%> ~ <%=rs_q("enddate")%>)</td>
        </tr>
        <tbody> 
        <tr bgcolor=#f5f5f5> 
          <td width="24%" bgcolor=#fff4e5 align="left"> ע���Ա��ʱ��</td>
          <td width="61%" bgcolor=#F7F7F7><%=rs_q("idate")%></td>
        </tr>
        <tr bgcolor=#f5f5f5> 
          <td bgcolor=#fff4e5> �����¼��ʱ��</td>
          <td bgcolor=#F7F7F7><%=rs_q("LastLoginTime")%></td>
        </tr>
        <tr bgcolor=#f5f5f5> 
          <td bgcolor=#fff4e5 align="left" rowspan="2"> �ѷ�����������Ϣ����</td>
          <td bgcolor=#F7F7F7>�����̻���<font color=#ff0000>&nbsp; </font><font color=#ff0000><%=conn.execute("select count(*) from info where gsid='"&rs_q("user")&"'")(0)%></font></td>
        </tr>
        <tr bgcolor=#f5f5f5> 
          <td bgcolor=#F7F7F7>չʾ��Ʒ��<font color=#ff0000> 
            <%set rs_p=conn.execute("select count(*) from spzs where gsid='"&rs_q("user")&"'")
				  response.write rs_p(0)
				  rs_p.close
				  set rs_p=nothing%>
            </font></td>
        </tr>
        </tbody> 
      </table>
      <table class=style height=30 cellspacing=0 cellpadding=0 width="96%" 
            align=center bgcolor=#fafafa border=0>
        <tbody> 
        <tr> 
          <td bgcolor=#FFCC00><span 
            class=zi1>��֤�鼰����</span></td>
        </tr>
        </tbody> 
      </table>
        <table class=style cellspacing=2 cellpadding=2 width="96%" 
            align=center border=0>
          <tbody>
            <tr bgcolor=#f5f5f5> 
              <td width="49%" height="89" align=middle bgcolor="#f7f7f7"> 
                <%set QyzsRs=Conn.Execute("Select * from qyzs where gsid='"&rs_q("user")&"'")
				   if QyzsRs.eof and QyzsRs.bof then
				       response.write "�ù�˾��δ������ҵ֤��!"
				   else
				   %>
                <table class=style cellspacing=3 cellpadding=3 width="100%" 
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
        <table class=style cellspacing=0 cellpadding=4 width="96%" 
            align=center border=0>
        <tbody> 
        <tr> 
          <td 
                  bgcolor=#fff4e5>��ע�⣺<br>
            ��������������ʾ������Ϣ���������������������յ������������֤��ʹ��������Ϣʱ�������֮�����Ҳ����������Ϣ��׼ȷ�ԣ������Ժͼ�ʱ�ԣ�Ҳ���е�����ߵ��κ���ҵ���ա���ˣ��������������е�����˶���������ʹ���κ��𺦡���������������ȫ���򲿷�ɾ�����������Ȩ��������ҳ�����ݽ�ֹ�Ը�����ʽ��ȫ���򲿷ָ���ʹ�á� 
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td valign=top width=15 bgcolor=#fcf6e8> 
      <table height=69 cellspacing=0 cellpadding=0 width="100%" 
      background=images/img_bg_9.gif border=0>
        <tbody> 
        <tr> 
          <td>&nbsp;</td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<!--#include file="Copyright.htm"-->
</BODY></HTML>
<%end if%>
