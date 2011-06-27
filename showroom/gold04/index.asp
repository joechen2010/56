<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/function.asp"-->
<%
Dim GsID
GsID=SafeRequest("GsID",1)
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
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 首页</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="FFFCF4">
  <tr>
    <td width="175">&nbsp;</td>
    <td width="1" background="images/hot_line.jpg"></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td align="center" valign="top"> 
      <!--#include file="left.htm"-->
	  </td>
    <td width="1" background="images/hot_line.jpg"></td>
    <td align="center" valign="top"> 
      <table width="96%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="images/skin3_15.jpg" width="167" height="27" align="center"><a class=white 
                        href="aboutus.asp?id=<%=gsid%>"><font 
                        color=#ffffff><b>产品展示</b></font></a></td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="96%" border="0" cellspacing="1" cellpadding="0" bgcolor="f37600">
        <tr> 
          <td bgcolor="f37600" height="3"></td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF"> 
            <table height=96 cellspacing=0 cellpadding=0 width="100%" border=0>
              <tbody> 
              <tr> 
                <td height=96><br>
                  <table cellspacing=0 cellpadding=0 width="100%" 
            align=center border=0>
                    <tbody> 
                    <tr> 
                      <%set rs_p=server.createobject("adodb.recordset")
			  sql="select * from spzs where flag=1 and gsid='"&rs_q("user")&"' Order By AddTime Desc"
			  rs_p.open sql,conn,1,1
			  if rs_p.eof and rs_p.bof then
			      response.write "<td>暂无产品发布！</td>"  
			  else
			  dim i
			  i=0
			    do while not rs_p.eof
			  %>
                      <td> 
                        <table cellspacing=2 cellpadding=2 width="100%" align=center 
                  border=0>
                          <tbody> 
                          <tr> 
                            <td align=middle> 
                              <%if rs_p("ProIMG")<>"" then%>
                              <a href="../../Product/ProDetail.asp?gsid=<%=gsid%>&ProID=<%=rs_p("ProID")%>" target="_blank"> 
                              <img  src="../../ProUpload/<%=rs_p("ProIMG")%>" width=75 border=1 height=75 style="border: 1px solid #000000"></a> 
                              <%else%>
                              <img src="../../Product/images/nopic.gif" width="75" height="75" border=1 style="border: 1px solid #000000"> 
                              <%end if%>
                            </td>
                          </tr>
                          <tr> 
                            <td class=style> 
                              <div align=center><a href="../../Product/ProDetail.asp?gsid=<%=gsid%>&ProID=<%=rs_p("ProID")%>" target="_blank"><%=rs_p("BearinNo")%></a></div>
                            </td>
                          </tr>
                          </tbody> 
                        </table>
                      </td>
                      <%
						i=i+1
	                     if i>=4 then exit do
	                    rs_p.movenext
	                   loop
					   end if 
					   rs_p.close
					   set rs_p=nothing
						%>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
      </table>
      <TABLE height=96 cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY> 
        <TR> 
          <TD height=96 align="center"> 
            <table width="96%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td background="images/skin3_15.jpg" width="167" height="27" align="center"><a class=white 
                        href="aboutus.asp?id=<%=gsid%>"><font 
                        color=#ffffff><b>公司简介</b></font></a></td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <table width="96%" border="0" cellspacing="1" cellpadding="0" bgcolor="f37600">
              <tr> 
                <td bgcolor="f37600" height="3"></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF"> 
                  <table height=96 cellspacing=0 cellpadding=0 width="100%" border=0>
                    <tbody> 
                    <tr> 
                      <td height=96>
                        <TABLE class=style cellSpacing=5 cellPadding=5 
                        width="100%" align=center border=0>
                          <TBODY> 
                          <TR> 
                            <TD height=85> 
                              <P> 
                                <%if rs_q("qyjj")<>"" then 
        response.write Replace(rs_q("qyjj"),chr(13),"<br>&nbsp;&nbsp;&nbsp;")
	else
	response.write rs_q("qyjj")
end if
%>
                              </P>
                            </TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                      </td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
            </table>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE height=96 cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY> 
        <TR> 
          <TD align="center"><BR>
            <table width="96%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td background="images/skin3_15.jpg" width="167" height="27" align="center"><a class=white 
                        href="aboutus.asp?id=<%=gsid%>"><font 
                        color=#ffffff><b>供求信息</b></font></a></td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <table width="96%" border="0" cellspacing="1" cellpadding="0" bgcolor="f37600">
              <tr> 
                <td bgcolor="f37600" height="3"></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF"> 
                  <table height=96 cellspacing=0 cellpadding=0 width="100%" border=0>
                    <tbody> 
                    <tr> 
                      <td height=96> 
                        <table height=43 cellspacing=2 cellpadding=0 width="100%" 
            align=center border=0>
                          <tbody> 
                          <tr> 
                <%set rs_info=server.createobject("adodb.recordset")
			  sql="select * from info where flag=1 and gsid='"&rs_q("user")&"' Order By AddTime Desc"
			  rs_info.open sql,conn,1,1
			  if rs_info.eof and rs_info.bof then
			      response.write "<td>暂无信息！</td>"
			  else
			  dim n
			  n=0
			  do while not rs_info.eof
			  %>
                <td width="49%"> 
                              <table class=style height=80 cellspacing=0 cellpadding=0 
                  width="100%" align=center border=0>
                                <tbody> 
                                <tr> 
                      <td width="20%"> 
                        <div align=center>【<font 
color=#ff0000><%=rs_info("infotype")%></font>】</div>
                      </td>
                      <td width="63%"><a href="../../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>"><%=rs_info("Title")%></a> 
                        <span class=zi2>(<%=rs_info("AddTime")%>)</span> <br>
                        <a href="../../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>">[<font color="#FF6600">详细信息</font>]</a><br>
                      </td>
                      <td width="17%"> 
                        <div align=center><font color=#000000><img height=32 
                        src="images/icon.gif" 
                        width=38><br>
                          <a 
                        href="message.asp?gsid=<%=gsid%>">询价</a></font></div>
                      </td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
                <%
			  n=n+1
	          if n>=4 then exit do
			  if n mod 2 =0 then
			     response.write "</tr><TR bgColor=#eef0fd>"
			  end if
	          rs_info.movenext
	          loop
			  end if
			  rs_info.close
			  set rs_info=nothing
			  %>
              </tr>
              </tbody> 
            </table>
            </td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
            </table>
            <BR>
          </TD>
        </TR>
        </TBODY>
      </TABLE>
    </td>
  </tr>
</table>
        


<!--#include file="Copyright.htm"-->
</BODY></HTML>

<%END IF%>