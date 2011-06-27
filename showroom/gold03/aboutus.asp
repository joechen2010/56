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
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 公司介绍</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
</HEAD>
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
                    width="95%">公司介绍</TD>
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
                <TD vAlign=top align=middle> 
                  <table class=style cellspacing=5 cellpadding=0 width="95%" 
                  align=center border=0>
                    <tbody> 
                    <tr> 
                      <td class=S2 align=middle bgcolor=#f9f9f9></td>
                    </tr>
                    <tr> 
                      <td class=S2 bgcolor=#f9f9f9> 
                        <p>&nbsp; </p>
                        <table align=right border=0 cellpadding=4 cellspacing=4>
                          <tbody> 
                          <tr> 
                            <td> 
                              <%if rs_q("cimg")<>"" then
		     response.write "<img src=""../../Member/"&rs_q("cimg")&""" height=180>"
		  else%>
                              <img 
            src="images/nopicture.gif" width=260> 
                              <%end if%>
                            </td>
                          </tr>
                          </tbody> 
                        </table>
                        <p>
                          <%if rs_q("qyjj")<>"" then 
        response.write Replace(rs_q("qyjj"),chr(13),"<br>&nbsp;&nbsp;&nbsp;")
	else
	response.write rs_q("qyjj")
end if
%>
                        </p>
                      </td>
                    </tr>
                    <tr> 
                      <td class=S> 
                        <table height=100 cellspacing=1 cellpadding=1 
                        width="100%" align=center bgcolor=#ffcccc border=0>
                          <tbody> 
                          <tr> 
                            <td bgcolor=#ffffff> 
                              <table cellspacing=0 cellpadding=5 width="100%" 
                              border=0>
                                <tbody> 
                                <tr> 
                                  <td bgcolor=#feedd1 height=100> 
                                    <table class=style cellspacing=5 cellpadding=5 
                                width="100%" align=center bgcolor=#ffffff 
                                border=0>
                                      <tbody> 
                                      <tr> 
                                        <td height=85> 
                                          <table cellspacing=2 cellpadding=3 width="100%" align=center border=0>
                                            <tbody> 
                                            <tr> 
                                              <td valign=top align=left width="50%" bgcolor=#fff4e5>主营产品或服务： 
                                                <br>
                                                <span class=other><%=rs_q("Zycp")%></span></td>
                                              <td valign=top align=left width="50%" bgcolor=#fff4e5>企业经营性质： 
                                                <br>
                                                <span class=other><%=rs_q("qyxz")%> 
                                                </span></td>
                                            </tr>
                                            <tr> 
                                              <td background="" colspan=2 height=1></td>
                                            </tr>
                                            <tr> 
                                              <td valign=top align=left width="50%">经营模式： 
                                                <br>
                                                <span 
            class=other><%=rs_q("Qylb")%></span></td>
                                              <td valign=top align=left>法人代表/企业负责人： 
                                                <br>
                                                <span class=other> <%=rs_q("Frdb")%></span></td>
                                            </tr>
                                            <tr> 
                                              <td background="" colspan=2 height=1></td>
                                            </tr>
                                            <tr> 
                                              <td valign=top align=left width="50%" bgcolor=#fff4e5>公司注册地： 
                                                <br>
                                                <span class=other><%=rs_q("province")%><%=rs_q("city")%> 
                                                </span></td>
                                              <td valign=top align=left bgcolor=#fff4e5>公司经营地： 
                                                <br>
                                                <span 
            class=other><%=rs_q("address")%> </span></td>
                                            </tr>
                                            <tr> 
                                              <td background="" colspan=2 height=1></td>
                                            </tr>
                                            <tr> 
                                              <td valign=top width="50%">员工人数： 
                                                <br>
                                                <span class=other><%=rs_q("Ygrs")%></span></td>
                                              <td valign=top align=left>注册资金： 
                                                <br>
                                                <span class=other> <%=rs_q("Zczj")%> 
                                                万元</span></td>
                                            </tr>
                                            <tr> 
                                              <td background="" colspan=2 height=1></td>
                                            </tr>
                                            </tbody> 
                                          </table>
                                        </td>
                                      </tr>
                                      </tbody> 
                                    </table>
                                  </td>
                                </tr>
                                </tbody> 
                              </table>
                            </td>
                          </tr>
                          </tbody> 
                        </table>
                      </td>
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

<!--#include file="Copyright.htm"-->
<%end if%>