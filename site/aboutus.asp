
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%infoid=request("infoid")
  user=request("user")
%>
		<%
        dim infoid
         infoid=Request.QueryString("infoid")
        if isnumeric(infoid)=0 or infoid="" then
         response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	     response.end
        end if  		
		set rs=server.CreateObject("adodb.recordset")
		  sql="select * from qyml where id="&request("infoid")&""
		  'sql="select * from info2 where infoid="&infoid
		  rs.open sql,conn,1,1
		%>
<HTML>
<HEAD>
<TITLE><%=rs("qymc")%>-诚信物流网</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>
BODY {
	MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px
}
.style1 {
	FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #ffffff
}
.style2 {
	FONT-SIZE: 12px; LINE-HEIGHT: 23px
}
.style3 {
	FONT-SIZE: 12px; LINE-HEIGHT: 17px
}
.style4 {
	FONT-WEIGHT: bold; FONT-SIZE: 14px
}
.style5 {
	FONT-SIZE: 28px
}
A.a03:link {
	FONT-SIZE: 12px; COLOR: #333333; TEXT-DECORATION: none
}
A.a03:visited {
	FONT-SIZE: 12px; COLOR: #333333; TEXT-DECORATION: none
}
A.a03:hover {
	FONT-SIZE: 12px; COLOR: #f46624; TEXT-DECORATION: underline
}
A.a04:link {
	FONT-SIZE: 12px; COLOR: #ffffff; TEXT-DECORATION: none
}
A.a04:visited {
	FONT-SIZE: 12px; COLOR: #ffffff; TEXT-DECORATION: none
}
A.a04:hover {
	FONT-SIZE: 12px; COLOR: #ffffff; TEXT-DECORATION: underline
}
A:link {
	FONT-SIZE: 12px; COLOR: #333333; TEXT-DECORATION: none
}
A:visited {
	FONT-SIZE: 12px; COLOR: #333333; TEXT-DECORATION: none
}
A:hover {
	FONT-SIZE: 12px; COLOR: #f4560d; TEXT-DECORATION: underline
}
A:active {
	FONT-SIZE: 12px; COLOR: #333333; TEXT-DECORATION: none
}
.style6 {
	COLOR: #000000
}
.style81 {
	FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 20px; FONT-FAMILY: "宋体"
}
.unnamed11 {
	FONT-SIZE: 13px; FONT-FAMILY: "宋体"
}
.style7 {
	FONT-SIZE: 12px; LINE-HEIGHT: 25px
}
BODY {
	FONT-SIZE: 12px
}
TD {
	FONT-SIZE: 12px
}
TH {
	FONT-SIZE: 12px
}
.unnamed111 {
	FONT-SIZE: 13px; LINE-HEIGHT: 23px; FONT-FAMILY: "宋体"
}
.STYLE82 {
	FONT-SIZE: 9px
}
</STYLE>

<META content="MSHTML 6.00.2800.1561" name=GENERATOR></HEAD>
<BODY>
<!--#include file="top.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TR> 
    <TD colSpan=2 height=120 background="images/topbg.jpg"> 
      <DIV class="style1 style5" align=left>&nbsp;&nbsp;
	  <%if not isnull(rs("logo")) then response.Write "<img src='../Member/"&rs("logo")&"'>"%>
	  &nbsp;<%=rs("qymc")%></DIV>
    </TD>
  </TR>
  <TR> 
    <TD width="19%"> 
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TR> 
          <TD vAlign=top width="17%" bgColor=#e8e8e8> 
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TR> 
                <TD><!--#include file="left.asp"--></TD>
              </TR>
            </TABLE>
          </TD>
          <TD vAlign=top width="83%"> 
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TR> 
                <TD class=style2 bgColor=#e8e8e8 height=20> 
                  <DIV align=left>&nbsp; 您现在的位置：<%=rs("qymc")%> &nbsp;&gt;<strong> 
                    公司简介</strong></DIV>
                </TD>
              </TR>
              <TR> 
                <TD height=520 bgcolor="#fde7db" valign="top"> 
                  <table cellspacing=0 cellpadding=0 width="97%" align=center border=0>
                    <tbody> 
                    <tr> 
                      <td class=righttitlel width=14 background=images/img_subject1_14x30.gif>&nbsp;</td>
                      <td class=righttitlem background=images/img_title_bg_1x30_.gif height=30><b>公司简介</b></td>
                      <td class=righttitler width=14 background=images/img_subject2_14x30.gif>&nbsp;</td>
                    </tr>
                    </tbody> 
                  </table>
                  <table height=64 cellspacing=0 cellpadding=0 width="97%" align=center border=0>
                    <tbody> 
                    <tr> 
                      <td style="PADDING-RIGHT: 5px; PADDING-LEFT: 15px; PADDING-BOTTOM: 8px; PADDING-TOP: 8px" valign=top bgcolor=#ffffff height=100> 
                        <%'if session("user")<>"" then%>
                        <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
                          <TR> 
                            <TD> 
                              <P class=style2>&nbsp;&nbsp;&nbsp;&nbsp;<%=dvHTMLEncode(rs("qyjj"))%>                              </P>                            </TD>
							<td><%if not isnull(rs("cimg")) then response.Write "<img src='../Member/"&rs("cimg")&"'>"%></td>
                          </TR>
                          <TR> 
                            <TD colspan="2">
<%
sql2="select zsurl from qyzs where gsid='"+request("user")+"' order by ID desc"
set rs2= Server.CreateObject("ADODB.Recordset") 
rs2.open sql2,conn,1,1 
if not (rs2.eof and rs2.bof) then
  k=0
 do while not rs2.eof 
 response.Write "<img src='../Member/"&rs2("zsurl")&"'>&nbsp;&nbsp;&nbsp;&nbsp;"
 k=k+1
 if k mod 3 =0 then response.Write "<br>"
 rs2.movenext
 loop
end if
%>
							</TD>
							</TR>						  
                        </TABLE>
                        <%'end if%>
                        <%'if session("user")="" then%>
                        <!--<TABLE cellSpacing=0 cellPadding=5 width=778 align=center border=0>
        <TR>
          <TD>   
   <%' Response.Write("<font color='#999999'>为保证信息发布者的隐私权，您需要<a href='../Login/login.asp' target='_blank'>登录</a>才可以查看联系方式</font>")
    'Response.Write("<br><a href='../Reg/User_Reg.asp' target='_blank'>未注册用户请点击>></a>")
	%>
		  </TD>
		</TR>
</TABLE>-->
                        <%'end if%>
                      </td>
                    </tr>
                    </tbody> 
                  </table>
                  <br>
                  <table cellspacing=0 cellpadding=0 width="97%" align=center border=0>
                    <tbody> 
                    <tr> 
                      <td class=righttitlel width=14 
          background=images/img_subject1_14x30.gif>&nbsp;</td>
                      <td class=righttitlem 
          background=images/img_title_bg_1x30_.gif 
          height=30><b>联系我们</b></td>
                      <td class=righttitler width=14 
          background=images/img_subject2_14x30.gif>&nbsp;</td>
                    </tr>
                    </tbody> 
                  </table>
                  <table class=table-border-product-detail cellspacing=0 cellpadding=4 
      width="97%" align=center border=0>
                    <tbody> 
                    <tr> 
                      <td bgcolor=#ffffff> 
                        <table cellspacing=0 cellpadding=0 width="100%" bgcolor=#ffffff 
            border=0>
                          <tbody> 
                          <tr> 
                            <td valign=center height=180> 
                              <table cellspacing=0 cellpadding=0 width="92%" align=center border=0>
                                <tr> 
                                  <td class=style2 height=120> 
                                    <%'if session("user")<>"" then%>
                                    <table cellspacing=1 cellpadding=4 width="100%" border=0 bgcolor="#FBCDB5">
                                      <tr bgcolor="#fde7db"> 
                                        <td width="10%"  align="right">所在城市：</td>
                                        <td width="30%"><%=rs("province")%>&nbsp;<%=rs("city")%>&nbsp;<%=rs("area")%></td>
                                        <td  align="right" width="10%">公司地址：</td>
                                        <td width="50%"><%=rs("address")%></td>
                                      </tr>
                                      <tr bgcolor="#FFFFFF"> 
                                        <td  align="right" width="10%">电　　话：</td>
                                        <td width="30%"><%=rs("phone2")%>-<%=rs("phone3")%></td>
                                        <td  align="right" width="10%">传　　真：</td>
                                        <td width="50%"><%=rs("fax2")%>-<%=rs("fax3")%></td>
                                      </tr>
                                      <tr bgcolor="#fde7db"> 
                                        <td  align="right" width="10%">联 系 人：</td>
                                        <td width="30%"><%=rs("name")%></td>
                                        <td  align="right" width="10%">手　　机：</td>
                                        <td width="50%"> 
                                          <%if session("user")<>"" then
										response.write rs("Mobile")
										else
										 response.write "*********** (请<a href='../Login/login.asp?action=login'>登陆</a>后查看)"
										end if
										%>
                                        </td>
                                      </tr>
                                      <tr bgcolor="#FFFFFF"> 
                                        <td  align="right" width="10%">E-mail：</td>
                                        <td width="30%"><%=rs("email")%></td>
                                        <td  align="right" width="10%">网　　址：</td>
                                        <td width="50%"><a href="<%=rs("web")%>" target="_blank"><%=rs("web")%></a></td>
                                      </tr>
                                      <tr bgcolor="#fde7db"> 
                                        <td  align="right" width="10%">邮　　编：</td>
                                        <td width="30%" bgcolor="#fde7db"><%=rs("post")%></td>
                                        <td  align="right" width="10%">注册时间：</td>
                                        <td width="50%"><%=rs("idate")%></td>
                                      </tr>
                                    </table>
                                    <%'end if%>
                                    <%'if session("user")="" then%>
                                    <!--<TABLE cellSpacing=0 cellPadding=5 width=778 align=center border=0>
        <TR>
          <TD>   
   <% 'Response.Write("<font color='#999999'>为保证信息发布者的隐私权，您需要<a href='../Login/login.asp' target='_blank'>登录</a>才可以查看联系方式</font>")
    'Response.Write("<br><a href='../Reg/User_Reg.asp' target='_blank'>未注册用户请点击>></a>")
	%>
		  </TD>
		</TR>
</TABLE>-->
                                    <%'end if%>
                                  </td>
                                </tr>
                                <tr> 
                                  <td height=22 align="center"> 
                                    <%if session("user")<>"" then%>
                                    <%if request("user")<>session("user") then%>
                                    <a href="javascript:openwindow('tj.asp?T_User=<%=request("user")%>',800,600)"> 
                                    <img height=37 src="images/208.gif" width=138 border=0></a> 
                                    <%end if%>
                                    <%end if%>
                                  </td>
                                </tr>
                              </table>
                              <font face=黑体 
                size=2></font></td>
                          </tr>
                          </tbody> 
                        </table>
                      </td>
                    </tr>
                    </tbody> 
                  </table>
                  <br>
                  <table bordercolor=#FFBD6E width="97%" align="center"  border=1>
                    <tbody> 
                    <tr > 
                      <td bgcolor="#FFFFFF"> 
                          
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td rowspan="2" height="13" width="4%"><img height=22 src="images/209.gif" 
                        width=23></td>
                            <td width="96%" height="22"><font color="#FF0000">免责声明：以上所展示的信息由企业自行提供，内容的真实性、准确性和合法性由发布企业负责。诚信物流网对此不承担任何保证责任。</font></td>
                          </tr>
                          <tr> 
                            <td width="96%" height="22"><font color="#FF0000">友情提醒：为保障您的利益， 
                              建议优先选择诚信物流网会员！ </font></td>
                          </tr>
                        </table>
                        </td>
                    </tr>
                    </tbody> 
                  </table>
                  
                </TD>
              </TR>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
  <TR> 
    <TD vAlign=bottom height=30> 
      <HR width="100%" noShade SIZE=1>
    </TD>
  </TR>
  <TR> 
    <TD vAlign=top><!--#include file="bottom.asp"--></TD>
  </TR>
</TABLE>
</BODY></HTML>
