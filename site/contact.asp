
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
<table cellspacing=0 cellpadding=0 width="100%" border=0>
  <tr> 
    <td colspan=2 height=120 background="images/topbg.jpg"><DIV class="style1 style5" align=left>&nbsp;&nbsp;
          <%if not isnull(rs("logo")) then response.Write "<img src='../Member/"&rs("logo")&"'>"%>
  &nbsp;<%=rs("qymc")%></DIV></td>
  </tr>
  <tr> 
    <td width="19%"> 
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
        <tr> 
          <td valign=top width="17%" bgcolor=#e8e8e8 height=185> 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
              <tr> 
                <td> 
                  <!--#include file="left.asp"-->
                </td>
              </tr>
            </table>
          </td>
          <td valign=top width="83%" bgcolor="#fde7db"> 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
              <tr> 
                <td class=style2 bgcolor=#e8e8e8 height=30> 
                  <div align=left>&nbsp; 您现在的位置：<%=rs("qymc")%> &nbsp;&gt;<strong> 
                    联系我们</strong></div>
                </td>
              </tr>
              <tr> 
                <td height=69 valign="top"> 
                  <table cellspacing=0 cellpadding=0 width="97%" align=center border=0>
                    <tbody> 
                    <tr> 
                      <td class=righttitlel width=14 
          background=images/img_subject1_14x30.gif>&nbsp;</td>
                      <td class=righttitlem 
          background=images/img_title_bg_1x30_.gif 
          height=30><strong>联系我们</strong></td>
                      <td class=righttitler width=14 
          background=images/img_subject2_14x30.gif>&nbsp;</td>
                    </tr>
                    </tbody> 
                  </table>
                  <table height=64 cellspacing=0 cellpadding=0 width="97%" align=center 
      border=0>
                    <tr> 
                      <td 
          style="PADDING-RIGHT: 5px; PADDING-LEFT: 15px; PADDING-BOTTOM: 8px; PADDING-TOP: 8px" 
          valign=top bgcolor=#ffffff height=160> 
                        <table cellspacing=0 cellpadding=0 width="75%" align=center border=0>
                          <tr> 
                            <td width="1" height=1 class=style2 bgcolor="#D0D0D0"></td>
                            <td width="562" height=1 class=style2 bgcolor="#D0D0D0"></td>
                            <td width="11" background=images/205.gif rowspan=3></td>
                          </tr>
                          <tr> 
                            <td rowspan="2" height=196 class=style2 width="1" bgcolor="#D0D0D0"></td>
                            <td width="562" height=174 class=style2> 
                              <%'if session("user")<>"" then%>
                              <table cellspacing=1 cellpadding=6 width="98%" border=0 bgcolor="#D9D9D9" align="center">
                                <tr bgcolor="#FFFFFF"> 
                                  <td width="15%"  align="right">所在城市：</td>
                                  <td colspan="3"><%=rs("province")%>&nbsp;<%=rs("city")%>&nbsp;<%=rs("area")%></td>
                                </tr>
                                <tr bgcolor="#E8E8E8"> 
                                  <td  align="right" width="15%">公司地址：</td>
                                  <td colspan="3"><%=rs("address")%></td>
                                </tr>
                                <tr bgcolor="#FFFFFF"> 
                                  <td  align="right" width="15%">电　　话：</td>
                                  <td width="35%" bgcolor="#FFFFFF"><%=rs("phone2")%>-<%=rs("phone3")%></td>
                                  <td  align="right" width="15%">传　　真：</td>
                                  <td width="35%"><%=rs("fax2")%>-<%=rs("fax3")%></td>
                                </tr>
                                <tr bgcolor="#E8E8E8"> 
                                  <td  align="right" width="15%">联 系 人：</td>
                                  <td width="35%"><%=rs("name")%></td>
                                  <td  align="right" width="15%">手　　机：</td>
                                  <td width="35%">
								  <%if session("user")<>"" then
										response.write rs("Mobile")
										else
										 response.write "*********** (请<a href='../Login/login.asp'>登陆</a>后查看)"
										end if
										%>
</td>
                                </tr>
                                <tr bgcolor="#FFFFFF"> 
                                  <td  align="right" width="15%">E-mail：</td>
                                  <td width="35%"><%=rs("email")%></td>
                                  <td  align="right" width="15%">网　　址：</td>
                                  <td width="35%"><a href="<%=rs("web")%>" target="_blank"><%=rs("web")%></a></td>
                                </tr>
                                <tr bgcolor="#E8E8E8"> 
                                  <td  align="right" width="15%">邮　　编：</td>
                                  <td width="35%"><%=rs("post")%></td>
                                  <td  align="right" width="15%">注册时间：</td>
                                  <td width="35%"><%=rs("idate")%></td>
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
                            <td height=22 align="center" width="562"> 
                              <%if session("user")<>"" then%>
                              <%if request("user")<>session("user") then%>
                              <a href="javascript:openwindow('tj.asp?T_User=<%=request("user")%>',800,600)"> 
                              <img height=37 src="images/208.gif" width=138 border=0></a> 
                              <%end if%>
                              <%end if%>
                            </td>
                          </tr>
                          <tr> 
                            <td background=images/206.gif height=8 width="1"></td>
                            <td background=images/206.gif height=8 width="562"><img 
                              height=1 src="images/spacer.gif" width=1></td>
                            <td width="11" background=images/205.gif><img 
                              height=9 src="images/207.gif" 
                          width=10></td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tbody> </tbody> 
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
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td valign=bottom height=30> 
      <hr width="100%" noShade size=1>
    </td>
  </tr>
  <tr> 
    <td valign=top> 
      <!--#include file="bottom.asp"-->
    </td>
  </tr>
</table>
</BODY></HTML>
