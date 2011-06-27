<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/function.asp"-->
<%
Dim GsID
GsID=SafeRequest("GsID",1)
if GsID="" then
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
                    width="95%">主推产品</TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                </TD>
                <TD><IMG height=32 src="images/menuri.gif" 
              width=11></TD>
              </TR>
              </TBODY> 
            </TABLE>
            <TABLE class=btBorder cellSpacing=0 cellPadding=0 width="100%" 
            bgColor=#ffffff>
              <TBODY> 
              <TR> 
                <TD vAlign=top align=middle> 
                  <table height=37 cellspacing=0 cellpadding=0 width="90%" 
            align=center background=images/title.jpg border=0>
                    <tbody> 
                    <tr> 
                      <td width="6%"><img height=32 src="images/pic.jpg" 
                  width=38></td>
                      <td width="22%"> 
                        <table height=24 cellspacing=0 cellpadding=0 width="69%" 
                  align=center bgcolor=#ffffff border=0>
                          <tbody> 
                          <tr> 
                            <td> 
                              <div class=tile2 
                      align=center><strong>产品展示</strong></div>
                            </td>
                          </tr>
                          </tbody> 
                        </table>
                      </td>
                      <td width="72%">&nbsp;</td>
                    </tr>
                    </tbody> 
                  </table>
                  <table height=43 cellspacing=0 cellpadding=0 width="90%" 
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
                        <table cellspacing=2 cellpadding=2 width="90%" align=center 
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
                </TD>
              </TR>
              </TBODY> 
            </TABLE>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="75%">
        <TBODY> 
        <TR> 
          <TD height=12></TD>
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
            <table class=style cellspacing=8 cellpadding=0 
                        width="100%" align=center bgcolor="#FFFFFF">
              <tbody> 
              <tr> 
                                  <td height=85> 
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
                                </tbody> 
                              </table>
                            </TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="75%">
        <TBODY> 
        <TR> 
          <TD height=12></TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE height=142 cellSpacing=0 cellPadding=0 width="97%">
        <TBODY> 
        <TR> 
          <TD vAlign=top align=middle bgcolor="#FFFFFF"> 
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
                    width="95%">最新供应</TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                </TD>
                <TD><IMG height=32 src="images/menuri.gif" 
              width=11></TD>
              </TR>
              </TBODY> 
            </TABLE>
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
                              <span class=zi2>(<%=rs_info("AddTime")%>)</span> 
                              <br>
                              <a href="../../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>">[<font color="#FF6600">详细信息</font>]</a><br>
                            </td>
                            <td width="17%"> 
                              <div align=center><font color=#000000><img height=32 
                        src="images/icon_offline.gif" 
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
                
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
      
    </TD>
  </TR>
  </TBODY> 
</TABLE>
<!--#include file="Copyright.htm"-->
</BODY></HTML>
<%END IF%>