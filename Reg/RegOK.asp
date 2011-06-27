<%
userid=request("userid")
pass=request("pass")
%>
<HTML>
<HEAD>

<TITLE>会员信息提交成功-诚信物流网</TITLE>
<META content=False name=vs_snapToGrid>
<META content=True name=vs_showGrid>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content=诚信物流网:全球最大的物流电子商务网站，提供物流品供求、物流品企业、物流产品、物流品项目、物流品新闻、物流品行情、物流品会展等信息的专业物流品网站。 
name=description>
<META content=全球最大的B2B物流平台、传承七千年物流文明精华，催生21世纪物流先锋、物流品供求、物流品、陶瓷物流、金属物流、泰蓝制品、艺术根雕、漆器雕漆、物流木雕、石雕物流、金箔制品、玉器玛瑙、红木摆件、奇石物流、象牙雕刻. 
name=keywords>
<META content=TRUE name=MSSmartTagsPreventParsing>
<META http-equiv=MSThemeCompatible content=Yes>
<LINK 
href="images/css.css" type=text/css rel=stylesheet>
<STYLE>.inputt01 {
	BORDER-RIGHT: rgb(139,139,139) 1px solid; BORDER-TOP: rgb(139,139,139) 1px solid; BORDER-LEFT: rgb(139,139,139) 1px solid; BORDER-BOTTOM: rgb(139,139,139) 1px solid
}
</STYLE>
<META content="MSHTML 6.00.2462.0" name=GENERATOR>
</HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE cellSpacing=0 cellPadding=0 width=776 align=center border=0>
  <TR> 
    <TD vAlign=top height=323> 
      <TABLE style="BORDER-BOTTOM: #5286b5 1px solid" height=70 cellSpacing=0 cellPadding=0 width=776 border=0>
        <TR> 
          <TD width=265 height=64><A href="../"><IMG height=65 
            alt=返回首页 src="images/logo.gif" width=265 
            align=bottom border=0></A></TD>
          <TD vAlign=bottom align=middle width=511> 
            <TABLE height=50 cellSpacing=0 cellPadding=0 width=489 border=0>
              <TR> 
                <TD width=180><IMG height=50 src="images/title_re1.gif" width=180></TD>
                <TD width=180><IMG height=50 src="images/title1_re2.gif" width=180></TD>
                <TD width=129><IMG height=50 src="images/title1_re3.gif" width=129></TD>
              </TR>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
        <TR> 
          <TD height=4></TD>
        </TR>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=776 border=0>
        <TR> 
          <TD> 
            <TABLE cellSpacing=0 cellPadding=0 width=100 border=0>
              <TR> 
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
              </TR>
            </TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TR> 
                <TD width=10>&nbsp;</TD>
                <TD align=left width=423><IMG height=68 src="images/regok.gif" width=423></TD>
                <TD align=right><IMG height=40 src="images/hotline.gif" width=230></TD>
              </TR>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=760 align=center border=0>
        <TR align=middle> 
          <TD width=760 height=40> 
            <TABLE cellSpacing=0 cellPadding=0 width=760 align=center border=0>
              <TR> 
                <TD vAlign=top align=middle width=427 bgColor=white> 
                  <TABLE cellSpacing=0 cellPadding=0 width=407 border=0>
                    <TR> 
                      <TD align=left> 
                        <P><IMG height=32 src="images/memberinfopic.gif" width=133></P>
                      </TD>
                    </TR>
                    <TR> 
                      <TD background=""> 
                        <P><IMG height=1 src="images/blank.gif" width=10 border=0></P>
                      </TD>
                    </TR>
                    <TR> 
                      <TD class=p14 align=center bgColor=#f7f7f7 height=50><SPAN class=meb1>&lt;您已经注册成功&gt;</SPAN> 
					  <br>请牢记您的登陆帐号和密码<br><FONT color=#000000>您的登陆帐号为:<%=userid%>&nbsp;&nbsp;密码为:<%=pass%></FONT>
                      </TD>
                    </TR>
                    <TR> 
                      <TD class=p14 align=left bgColor=#f7f7f7 height=5> 
                        <TABLE cellSpacing=0 cellPadding=0 width="98%" align=right border=0>
                          <TR> 
                            <TD bgColor=#cccccc height=1></TD>
                          </TR>
                          <TR> 
                            <TD bgColor=#ffffff height=1></TD>
                          </TR>
                        </TABLE>
                      </TD>
                    </TR>
                    <TR> 
                      <TD align=right bgColor=#f7f7f7 height=40><SPAN 
                        class=p141>想注册高级会员或特级会员</SPAN><A class=bai 
                        href="#">请点进这里</A> </TD>
                    </TR>
                    <TR> 
                      <TD background=""> 
                        <P><IMG height=1 src="images/blank.gif" 
                        width=10 border=0></P>
                      </TD>
                    </TR>
                  </TABLE>
                  </TD>
                <TD vAlign=top width=333 bgColor=white> 
                  <TABLE cellSpacing=1 borderColorDark=#ffffff cellPadding=0 width="100%" bgColor=#cab386 border=0>
                    <TR> 
                      <TD align=middle bgColor=#fcf5e0> 
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                          <TR> 
                            <TD align=left 
                            background=images/login_top_bg.gif><IMG 
                              height=25 src="images/login_top.gif" 
                              width=271></TD>
                          </TR>
                           
                        </TABLE>
                        <TABLE cellSpacing=5 cellPadding=0 width=300 
                          border=0>
                          <FORM name="MemberSubForm" action="../Login/Login_check.asp">
                            <TR> 
                              <TD class=S align=right width=60 height=25>会员帐号</TD>
                              <TD width=4 rowSpan=3>&nbsp;</TD>
                              <TD width=89> 
                                <INPUT class=inputt01 id=userid Size=14 value="<%=userid%>" name=UserID>
                              </TD>
                              <TD width=122 rowSpan=2> 
                                <INPUT id=ImageButton1 type=image alt="" src="images/login.gif" border=0 name=ImageButton1>
                              </TD>
                            </TR>
                            <TR> 
                              <TD class=S align=right height=25>会员密码</TD>
                              <TD> 
                                <INPUT class=inputt01 id=password3 type=password size=14 value=<%=pass%> name=Pwd>
                              </TD>
                            </TR>
                            <TR> 
                              <TD></TD>
                              <TD>&nbsp;</TD>
                              <TD></TD>
                            </TR>
                          </FORM>
                        </TABLE>
                      </TD>
                    </TR>
                  </TABLE>
                  </TD>
              </TR>
            </TABLE>
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>
<!--#include file="../inc/down.htm"-->
</BODY>
</HTML>
