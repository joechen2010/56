<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/function.asp"-->
<%
Dim gsID
gsID=SafeRequest("gsid",1)
if gsID="" then
     call WriteErrMsg2()
end if
set rs_q=server.createobject("adodb.recordset")
sql="select * from qyml where id="&gsid&""
rs_q.open sql,conn,1,1
if rs_q.eof and rs_q.bof then
     call WriteErrMsg2()
else
%>
<HTML>
<HEAD>
<TITLE><%=rs_q("qymc")%> - ����ͨ������ҵ - ��ϵ����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>

<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center>
  <TBODY> 
  <TR bgColor=#fc8410> 
    <TD colSpan=2 height=3></TD>
  </TR>
  <TR bgColor=#cf6903> 
    <TD colSpan=2 height=2></TD>
  </TR>
  <TR> 
    <TD vAlign=top align=middle width=156 bgColor=#ff822e>
<!--#include file="left.htm"-->      
	  </TD>
    <TD vAlign=top align=middle bgColor=#fff6f1> 
      <TABLE cellSpacing=0 cellPadding=0 width="75%">
        <TBODY> 
        <TR> 
          <TD height=8></TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE height=142 cellSpacing=0 cellPadding=0 width="97%">
        <TBODY> 
        <TR> 
          <TD vAlign=top align=middle> 
            <TABLE cellSpacing=0 cellPadding=0 width="100%">
              <TBODY> 
              <TR> 
                <TD height=32><IMG height=32 src="images/menule.gif" 
                  width=9></TD>
                <TD width="100%" background=images/menumi.gif> 
                  <TABLE cellSpacing=0 cellPadding=0 width="75%">
                    <TBODY> 
                    <TR> 
                      <TD width="5%"><IMG height=14 
                        src="images/iconrighl.gif" width=14></TD>
                      <TD class=font14 vAlign=bottom 
                    width="95%">��˾����</TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                </TD>
                <TD><IMG height=32 src="images/menuri.gif" 
              width=11></TD>
              </TR>
              </TBODY> 
            </TABLE>
            <TABLE class=btBorder cellSpacing=8 cellPadding=0 width="100%" 
            bgColor=#ffffff>
              <TBODY> 
              <TR> 
                <TD vAlign=top align=middle> <br>
                  <table class=style cellspacing=5 cellpadding=3 width="90%" 
            align=center border=0>
                    <tbody> 
                    <tr> 
                      <td class=C valign=top width="23%" bgcolor=#f3f3f3>��ϵ�ˣ�</td>
                      <td class=C width="77%" bgcolor=#f3f3f3><span class=C><%=rs_q("name")%> 
                        (<%=rs_q("ch")%>)&nbsp;<font size="-1"><b> 
                        <%if rs_q("zw")="��" then
else 
   response.write rs_q("zw")
end if%>
                        </b></font></span></td>
                    </tr>
                    <tr> 
                      <td class=C width="23%">�ֻ���</td>
                      <td class=C width="77%"><%=rs_q("mobile")%></td>
                    </tr>
                    <tr> 
                      <td class=C width="23%" bgcolor=#f3f3f3>�绰��</td>
                      <td class=C width="77%" bgcolor=#f3f3f3><font face="Arial">
                        <% Response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3")%>
                        </font></td>
                    </tr>
                    <tr> 
                      <td class=C width="23%">���棺</td>
                      <td class=C width="77%"><font face="Arial">
                        <% Response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
                        </font></td>
                    </tr>
                    <tr> 
                      <td class=C width="23%" bgcolor=#f3f3f3>��ַ��</td>
                      <td class=C width="77%" 
                  bgcolor=#f3f3f3><%=rs_q("address")%></td>
                    </tr>
                    <tr> 
                      <td class=C width="23%">�������룺</td>
                      <td class=C width="77%"><%=rs_q("post")%></td>
                    </tr>
                    <tr bgcolor="#f3f3f3"> 
                      <td class=C width="23%">��ַ��</td>
                      <td class=C 
width="77%"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a></td>
                    </tr>
                    </tbody> 
                  </table>
                </TD>
              </TR>
              </TBODY> 
            </TABLE>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
      
    </TD>
  </TR>
  </TBODY> 
</TABLE>

<%rs_q.close 
set rs_q=nothing
conn.close
set conn=nothing
%>
<%end if%>
<!--#include file="Copyright.htm"--></BODY></HTML>
