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
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 公司介绍</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file="head.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=189 bgColor=#fcf6e8> 
      <TABLE height=69 cellSpacing=0 cellPadding=0 width="100%" 
      background=images/img_bg_9.gif border=0>
        <TBODY> 
        <TR> 
          <TD valign="bottom"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="20"> 
                  <div align="center"><font color=#ff0000>物流通会员：创建于<%=rs_q("idate")%></font></div>
                </td>
              </tr>
            </table>
          </TD>
          <TD align=right width=10><IMG height=69 src="images/img_11.gif" 
            width=10></TD>
        </TR>
        </TBODY>
      </TABLE>
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
                <td width="22%" nowrap class=S2 height="28">地址：</td>
                <td width="78%" colspan="3" class=S2 height="28"><%=rs_q("address")%></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">邮编：</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"><%=rs_q("post")%></font></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">电话：</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"> 
                  <% Response.write rs_q("phone1")&"-"&rs_q("phone2")&"-"&rs_q("phone3")%>
                  </font></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">传真：</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"> 
                  <% Response.write rs_q("fax1")&"-"&rs_q("fax2")&"-"&rs_q("fax3")%>
                  </font></td>
              </tr>
              <tr> 
                <td class=S2 width="22%" height="28">主页：</td>
                <td class=S2 colspan="3" width="78%" height="28"><font face="Arial"><a href="<%=rs_q("web")%>"><%=rs_q("web")%></a>&nbsp; 
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </TD>
    <TD vAlign=top align=middle> 
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY> 
        <TR> 
          <TD vAlign=top height=12><IMG height=10 src="images/img_12.gif" 
            width=11></TD>
          <TD vAlign=top align=right><IMG height=10 src="images/img_14.gif" 
            width=10></TD>
        </TR>
        </TBODY>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="96%" 
      background=images/img_bg_31.gif border=0>
        <TBODY> 
        <TR> 
          <TD width="4%" height=29><IMG height=29 src="images/img_30.gif" 
            width=13></TD>
          <TD class=font14 width="92%"><STRONG>公司简介</STRONG></TD>
          <TD align=right width="4%"><IMG height=29 src="images/img_33.gif" 
            width=10></TD>
        </TR>
        </TBODY>
      </TABLE>
      <table height=200 cellspacing=0 cellpadding=0 width="96%" border=0>
        <tbody> 
        <tr> 
          <td> 
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
            <%if rs_q("qyjj")<>"" then 
        response.write Replace(rs_q("qyjj"),chr(13),"<br>&nbsp;&nbsp;&nbsp;")
	else
	response.write rs_q("qyjj")
end if
%>
          </td>
        </tr>
        </tbody> 
      </table>
      <TABLE cellSpacing=0 cellPadding=0 width="96%" 
      background=images/img_bg_31.gif border=0>
        <TBODY> 
        <TR> 
          <TD width="4%" height=29><IMG height=29 src="images/img_30.gif" 
            width=13></TD>
          <TD class=font14 width="92%"><STRONG>最新产品</STRONG></TD>
          <TD align=right width="4%"><IMG height=29 src="images/img_33.gif" 
            width=10></TD>
        </TR>
        </TBODY>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="96%" border=0>
        <TBODY> 
        <TR> 
          <TD> 
            <table height=96 cellspacing=0 cellpadding=0 width="100%" border=0>
              <tbody> 
              <tr> 
                <td>
<table height=43 cellspacing=0 cellpadding=0 width="98%" 
            align=center border=0>
                    <tr>
                      <td height="10"></td>
                    </tr>
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
                        <table cellspacing=2 cellpadding=2 width="95%" align=center 
                  border=0>
                          <tbody> 
                          <tr> 
                            <td align=middle> 
                              <%if rs_p("ProIMG")<>"" then%>
                              <a href="../../Product/ProDetail.asp?gsid=<%=gsid%>&ProID=<%=rs_p("ProID")%>" target="_blank"> 
                              <img  src="../../ProUpload/<%=rs_p("ProIMG")%>" width=72 border=1 height=72 style="border: 1px solid #000000"></a> 
                              <%else%>
                              <img src="../../Product/images/nopic.gif" width="72" height="72" border=1 style="border: 1px solid #000000"> 
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
          </TD>
        </TR>
        </TBODY>
      </TABLE>
      <TABLE height=27 cellSpacing=0 cellPadding=0 width="96%" 
      background=images/img_bg_37.gif border=0>
        <TBODY> 
        <TR> 
          <TD width="4%"><IMG height=27 src="images/img_36.gif" width=13></TD>
          <TD class=font14 width="92%"><STRONG>最新供求</STRONG></TD>
          <TD align=right width="4%"><IMG height=27 src="images/img_39.gif" 
            width=11></TD>
        </TR>
        </TBODY>
      </TABLE>
      <table height=43 cellspacing=2 cellpadding=0 width="97%" 
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
            <table class=style height=60 cellspacing=2 cellpadding=0 
                  width="100%" align=center border=0>
              <tbody> 
              <tr bgcolor="#F7F7F7"> 
                <td width="13%"> 
                  <div align=center>【<font 
color=#ff0000><%=rs_info("infotype")%></font>】</div>
                </td>
                <td width="70%"><a href="../../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>"><%=rs_info("Title")%></a> 
                  <span class=zi2>(<%=rs_info("AddTime")%>)</span> <a href="../../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>">[<font color="#FF0000">详细信息</font>]</a><br>
                </td>
                <td width="17%"> 
                  <div align=center><font color=#000000><img height=32 
                        src="images/icon.gif" 
                        width=38><br>
                    <a 
                        href="message.asp?Gsid=<%=gsid%>">询价</a></font></div>
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
    <TD vAlign=top width=15 bgColor=#fcf6e8>
      <TABLE height=69 cellSpacing=0 cellPadding=0 width="100%" 
      background=images/img_bg_9.gif border=0>
        <TBODY>
        <TR>
          <TD>&nbsp;</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<!--#include file="Copyright.htm"-->
</BODY></HTML>
<%END IF%>