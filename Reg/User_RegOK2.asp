<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim RegUserName,Password,confirmPassword,question,answer,sname,ch,Province,city,area,address,post,zw,phone,mobile,fax,email
dim qymc,web,qylb,sortid,jyfw,qyxz,qyjj,errmsg,founderr
RegUserName=request.form("user")
Password=request.form("pass")
confirmPassword=request.form("confirmPassword")
question=request.form("question")
answer=request.form("answer")
sname=request.form("name")
ch=request.form("ch")
Province=request.form("Province")
city=request.form("city")
area=request.form("area")
address=request.form("address")
post=request.form("post")
zw=request.form("zw")
phone1=request.form("phone1")
phone2=request("phone2")
phone3=request("phone3")
mobile=request.form("mobile")
fax1=request.form("fax1")
fax2=request.form("fax2")
fax3=request("fax3")
email=request.form("email")

qymc=request.form("qymc")
web=request.form("web")
if web="http://" then web=""
if len(web)>7 and left(web,7) <> "http://" then web="http://" & web
qylb=request.form("qylb")
sortid=request.form("sortid")
'typeid=request.form("typeid")
jyfw=request.form("jyfw")
Qyxz=request.form("Qyxz")
qyjj=request.form("qyjj")

if RegUserName="" or strLength(RegUserName)>14 or strLength(RegUserName)<4 then
	founderr=true
	errmsg=errmsg & "<br><li>请输入用户名(不能大于14小于4)</li>"
else
  	if Instr(RegUserName,"<")>0 or Instr(RegUserName,">")>0 or Instr(RegUserName,"=")>0 or Instr(RegUserName,"%")>0 or Instr(RegUserName,chr(32))>0 or Instr(RegUserName,"?")>0 or Instr(RegUserName,"&")>0 or Instr(RegUserName,";")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,"'")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,chr(34))>0 or Instr(RegUserName,chr(9))>0 or Instr(RegUserName,"")>0 or Instr(RegUserName,"$")>0 then
		errmsg=errmsg+"<br><li>用户名中含有非法字符</li>"
		founderr=true
	end if
end if
if Password="" or strLength(Password)>12 or strLength(Password)<6 then
	founderr=true
	errmsg=errmsg & "<br><li>请输入密码(不能大于12小于6)</li>"
else
	if Instr(Password,"<")>0 or Instr(Password,">")>0 or Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
		errmsg=errmsg+"<br><li>密码中含有非法字符</li>"
		founderr=true
	end if
end if
if ConfirmPassword="" then
	founderr=true
	errmsg=errmsg & "<br><li>请输入确认密码(不能大于12小于6)</li>"
else
	if Password<>confirmPassword then
		founderr=true
		errmsg=errmsg & "<br><li>密码和确认密码不一致</li>"
	end if
end if
if Question="" then
	founderr=true
	errmsg=errmsg & "<br><li>密码提示问题不能为空</li>"
end if
if Answer="" then
	founderr=true
	errmsg=errmsg & "<br><li>密码答案不能为空</li>"
end if

if qymc="" then
	founderr=true
	errmsg=errmsg & "<br><li> 公司名称不能为空</li>"
end if

if SortID="" then
	founderr=true
	errmsg=errmsg & "<br><li> 请选择产业种类！</li>"
end if

if qylb="" then
	founderr=true
	errmsg=errmsg & "<br><li> 请选择经营模式！</li>"
end if

if qyxz="" then
	founderr=true
	errmsg=errmsg & "<br><li> 请选择企业性质！</li>"
end if

if jyfw="" then
	founderr=true
	errmsg=errmsg & "<br><li> 请填写经营范围！</li>"
end if

if Email="" then
	founderr=true
	errmsg=errmsg & "<br><li>Email不能为空</li>"
else
	if IsValidEmail(Email)=false then
		errmsg=errmsg & "<br><li>您的Email有错误</li>"
   		founderr=true
	end if
end if

if qyjj="" then
	founderr=true
	errmsg=errmsg & "<br><li> 请填写企业简介！</li>"
end if

if founderr=false then
	dim sqlReg,rsReg,sqladminname,rsadminname

	sqladminname="select * from Manage_user where UserName='" & RegUserName & "'"
	set rsadminname=conn.execute(sqladminname)
	
	if not(rsadminname.bof and rsadminname.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>朋友，想冒充管理员？换一个用户名吧老大！</li>"
	else
	
    set rs=server.CreateObject ("adodb.recordset")
      sql="select * from qyml where user='"&RegUserName&"'"
      rs.open sql,conn,1,3
      if not (rs.eof and rs.bof) then
           founderr=true
		   errmsg=errmsg & "<br><li>你注册的用户已经存在！请换一个用户名再试试！</li>"
      else
          rs.addnew
          rs("user")=RegUserName
          rs("pass")=Password
          rs("question")=question
          rs("answer")=answer
          rs("name")=sname
          rs("ch")=ch
          rs("Province")=Province
          rs("city")=city
		  rs("area")=area
          rs("address")=address
          rs("post")=post
          rs("zw")=zw
          rs("phone1")=phone1
		  rs("phone2")=phone2
		  rs("phone3")=phone3
          rs("mobile")=mobile
          rs("fax1")=fax1
		  rs("fax2")=fax2
		  rs("fax3")=fax3
          rs("email")=email
          rs("qymc")=qymc
          rs("web")=web
          rs("qylb")=qylb
          rs("jygs")=jyfw
          rs("sortid")=sortid
          rs("idate")=date()
          rs("Flag")=0
          rs("uflag")=0
          rs("click")=0
          rs("Qyxz")=Qyxz
          rs("qyjj")=dvHTMLEncode(qyjj)
          rs.update
		  call SendRegEmail()
          rs.close
          set rs=nothing
	  end if
    end if
end if

              
if founderr=false then
	call RegSuccess()
else
	call WriteErrmsg()
end if

sub WriteErrMsg()
    response.write "<br><br><table align='center' width='300' border='0' cellpadding='5' bacellspacing='0' style='BACKGROUND-COLOR:#E4EEFD;BORDER:#FFFFFF 1PX SOLID;PADDING:0PX 6PS;'>"
    response.write "<tr class='title'><td align='center' height='22'>由于以下的原因不能注册用户！</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>" & errmsg & "<p align='center'>【<a href='javascript:onclick=history.go(-1)'>返 回</a>】<br></p></td></tr>"
	response.write "</table>" 
end sub

Sub SendRegEmail()
Set msg = Server.CreateObject("JMail.Message")
msg.silent = true
msg.Logging = true
msg.Charset = "gb2312"
msg.ContentType = "text/html"
msg.MailServerUserName = "webmaster@cxcnc.com" '输入smtp服务器验证登陆名 （邮局中任何一个用户的Email地址）
msg.MailServerPassword = "87679698"   '输入smtp服务器验证密码 (用户Email帐号对应的密码） 
msg.From = "webmaster@cxcnc.com"         '发件人Email
msg.FromName = "诚信物流网"     '发件人姓名
msg.AddRecipient email     '收件人Email
msg.Subject = "诚信物流网会员注册的通知"  '信件主题
msg.Body = "    感谢你在诚信物流网（<a href=""http://www.bearingpage.com"">http://www.bearingpage.com</a>）注册会员，你的用户名是："&RegUserName&",您的密码："&Password&"<br>"   
msg.Body = msg.Body &"-----------------------------------------------------<br>"
msg.Body = msg.Body &"联系电话：0574-63809664<br>"
msg.Body = msg.Body &"传真：0574-27860631<br>"
msg.Body = msg.Body &"地址：宁波市药行街31号F4-C5"
msg.Send ("smtp.cxcnc.com")              'smtp服务器地址（企业邮局地址）
set msg = nothing
end sub

sub RegSuccess()
dim i,MaxNum
MaxNum=5
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
  <TBODY> 
  <TR> 
    <TD vAlign=top height=323> 
      <TABLE style="BORDER-BOTTOM: #5286b5 1px solid" height=70 cellSpacing=0 
      cellPadding=0 width=776 border=0>
        <TBODY> 
        <TR> 
          <TD width=265 height=64><A href="../"><IMG height=65 
            alt=返回首页 src="images/logo.gif" width=265 
            align=bottom border=0></A></TD>
          <TD vAlign=bottom align=middle width=511> 
            <TABLE height=50 cellSpacing=0 cellPadding=0 width=489 border=0>
              <TBODY> 
              <TR> 
                <TD width=180><IMG height=50 
                  src="images/title_re1.gif" width=180></TD>
                <TD width=180><IMG height=50 
                  src="images/title1_re2.gif" width=180></TD>
                <TD width=129><IMG height=50 
                  src="images/title1_re3.gif" 
              width=129></TD>
              </TR>
              </TBODY> 
            </TABLE>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
        <TBODY> 
        <TR> 
          <TD height=4></TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=776 border=0>
        <TBODY> 
        <TR> 
          <TD> 
            <TABLE cellSpacing=0 cellPadding=0 width=100 border=0>
              <TBODY> 
              <TR> 
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
              </TR>
              </TBODY> 
            </TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY> 
              <TR> 
                <TD width=10>&nbsp;</TD>
                <TD align=left width=423><IMG height=68 
                  src="images/regok.gif" width=423></TD>
                <TD align=right><IMG height=40 
                  src="images/hotline.gif" 
            width=230></TD>
              </TR>
              </TBODY> 
            </TABLE>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=760 align=center border=0>
        <TBODY> 
        <TR align=middle> 
          <TD width=760 height=40> 
            <TABLE cellSpacing=0 cellPadding=0 width=760 align=center 
              border=0>
              <TBODY> 
              <TR> 
                <TD vAlign=top align=middle width=427 bgColor=white> 
                  <TABLE cellSpacing=0 cellPadding=0 width=407 border=0>
                    <TBODY> 
                    <TR> 
                      <TD align=left> 
                        <P><IMG height=32 
                        src="images/memberinfopic.gif" 
                        width=133></P>
                      </TD>
                    </TR>
                    <TR> 
                      <TD background=""> 
                        <P><IMG height=1 src="images/blank.gif" 
                        width=10 border=0></P>
                      </TD>
                    </TR>
                    <TR> 
                      <TD class=p14 align=left bgColor=#f7f7f7 height=50><FONT 
                        color=#000000>&nbsp;&nbsp;&nbsp;&nbsp;</FONT><SPAN 
                        class=S>您在诚信物流网注册的会员帐号已经成功激活,可以使用了!<BR>
                        &nbsp; 您当前注册的帐号类型为: </SPAN><SPAN class=meb1>&lt;普通会员&gt;</SPAN> 
                      </TD>
                    </TR>
                    <TR> 
                      <TD class=p14 align=left bgColor=#f7f7f7 height=5> 
                        <TABLE cellSpacing=0 cellPadding=0 width="98%" 
                        align=right border=0>
                          <TBODY> 
                          <TR> 
                            <TD bgColor=#cccccc height=1></TD>
                          </TR>
                          <TR> 
                            <TD bgColor=#ffffff 
                      height=1></TD>
                          </TR>
                          </TBODY> 
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
                    </TBODY> 
                  </TABLE>
                  <TABLE cellSpacing=0 cellPadding=0 width=407 border=0>
                    <TBODY> 
                    <TR> 
                      <TD align=left> 
                        <P><IMG height=32 
                        src="images/memberinfopic.gif" 
                        width=133></P>
                      </TD>
                    </TR>
                    <TR> 
                      <TD background=""> 
                        <P><IMG height=1 src="images/blank.gif" 
                        width=10 border=0></P>
                      </TD>
                    </TR>
                    <TR> 
                      <TD class=p141 align=left bgColor=#f7f7f7 
                        height=60><FONT 
                        color=#000000>&nbsp;&nbsp;&nbsp;&nbsp;</FONT><SPAN 
                        class=S>您在诚信物流网注册的会员帐号已经成功激活,可以使用了!<BR>
                        &nbsp; 您当前注册的帐号类型为: </SPAN><SPAN class=meb2>&lt;物流通会员&gt;</SPAN> 
                        <SPAN class=p12>(目前只能使用普通会员权限)</SPAN></TD>
                    </TR>
                    <TR> 
                      <TD class=p141 align=left bgColor=#f7f7f7 height=5> 
                        <TABLE cellSpacing=0 cellPadding=0 width="98%" 
                        align=right border=0>
                          <TBODY> 
                          <TR> 
                            <TD bgColor=#cccccc height=1></TD>
                          </TR>
                          <TR> 
                            <TD bgColor=#ffffff 
                      height=1></TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                      </TD>
                    </TR>
                    <TR> 
                      <TD class=p12 align=right bgColor=#f7f7f7 
                        height=40>我们将在三个工作日内与你联系核实会员真实信息并开通物流通会员服务!<BR>
                        <SPAN 
                        class=p14hon>你也可以拨打服务热线:0574-63809664联系</SPAN> </TD>
                    </TR>
                    <TR> 
                      <TD background=""> 
                        <P><IMG height=1 src="images/blank.gif" 
                        width=10 border=0></P>
                      </TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                  <TABLE cellSpacing=0 cellPadding=0 width=407 border=0>
                    <TBODY> 
                    <TR> 
                      <TD vAlign=bottom height=10> 
                        <HR align=center color=#e2e2e2 noShade SIZE=1>
                      </TD>
                    </TR>
                    <TR> 
                      <TD align=left> 
                        <TABLE cellSpacing=0 cellPadding=0 width=407 border=0>
                          <TBODY> 
                          <TR> 
                            <TD align=left height=30><B><FONT 
                              color=#ff6600>::</FONT> </B>您的帐号是: <SPAN 
                              id=userid2><%=RegUserName%> </SPAN><B><BR>
                              <FONT 
                              color=#ff6600>::</FONT> </B>您的密码是: <SPAN 
                              id=password2><%=Password%></SPAN> <B><BR>
                              </B> 
                              <HR align=center color=#e2e2e2 noShade SIZE=1>
                            </TD>
                          </TR>
                          <TR> 
                            <TD vAlign=top> 
                              <P>&nbsp;</P>
                            </TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=3 border=0>
                          <TBODY> 
                          <TR> 
                            <TD vAlign=top align=left width=400><SPAN 
                              class=S><IMG height=15 
                              src="images/icon4.gif" width=15> 请牢记以上会员注册信息,如果您忘记了诚信物流网会员登陆密码可以及时找回!</SPAN><BR>
                              <FONT 
                              color=#cc2800>∝ 也可以拨打诚信物流网客服热线:0574-63809664</FONT></TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                        <BR>
                        <HR align=center color=#e2e2e2 noShade SIZE=1>
                        <TABLE cellSpacing=0 cellPadding=0 width=407 border=0>
                          <TBODY> 
                          <TR> 
                            <TD align=left bgColor=#f2f2f2 height=25> 
                              <P><B><FONT color=#ff6600>::</FONT> 诚信物流网物流通会员基本介绍: 
                                <BR>
                                </B></P>
                            </TD>
                          </TR>
                          <TR> 
                            <TD bgColor=#999999 height=1></TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                        <P><SPAN class=p14hon>物流通会员基本服务（1800元/年）：</SPAN><BR>
                          <SPAN 
                        class=S>1. 随时发布浏览供应，求购等信息。</SPAN><SPAN 
                        class=p12><SPAN class=p14> <BR>
                          </SPAN></SPAN><SPAN 
                        class=S>2. 您发布的信息将予优先排名。 </SPAN><SPAN class=p12><SPAN 
                        class=p14><BR>
                          </SPAN></SPAN><SPAN class=S>3. 享用诚信物流网高级会员荣誉标志，大大提升您的企业品位和成交机率。</SPAN><SPAN 
                        class=p12><SPAN class=p14><BR>
                          </SPAN></SPAN><SPAN 
                        class=S>4. </SPAN><BR>
                          <BR>
                          <SPAN 
                        class=p14hon>VIP会员基本服务（6800元/年）：</SPAN><BR>
                          <SPAN 
                        class=S>1. 随时发布浏览供应，求购和合作等信息。 </SPAN><SPAN 
                        class=p12><SPAN class=p14><BR>
                          </SPAN></SPAN><SPAN 
                        class=S>2. 您发布的信息将予优先排名。</SPAN><SPAN class=p12><SPAN 
                        class=p14> <BR>
                          </SPAN></SPAN><SPAN class=S>3. 享用诚信物流网高级会员荣誉标志，大大提升您的企业品位和成交机率。</SPAN><SPAN 
                        class=p12><SPAN class=p14><BR>
                          </SPAN></SPAN><SPAN 
                        class=S>4. </SPAN></P>
                      </TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                </TD>
                <TD vAlign=top width=333 bgColor=white> 
                  <TABLE cellSpacing=1 borderColorDark=#ffffff cellPadding=0 
                  width="100%" bgColor=#cab386 border=0>
                    <TBODY> 
                    <TR> 
                      <TD align=middle bgColor=#fcf5e0> 
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY> 
                          <TR> 
                            <TD align=left 
                            background=images/login_top_bg.gif><IMG 
                              height=25 src="images/login_top.gif" 
                              width=271></TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                        <TABLE cellSpacing=5 cellPadding=0 width=300 
                          border=0>
                          <FORM name="MemberSubForm" action="../Login/Login_check.asp">
                            <TR> 
                              <TD class=S align=right width=60 height=25>会员帐号</TD>
                              <TD width=4 rowSpan=3>&nbsp;</TD>
                              <TD width=89> 
                                <INPUT class=inputt01 id=userid Size=14 value="<%=RegUserName%>" name=UserID>
                              </TD>
                              <TD width=122 rowSpan=2> 
                                <INPUT id=ImageButton1 
                              type=image alt="" 
                              src="images/login.gif" border=0 
                              name=ImageButton1>
                              </TD>
                            </TR>
                            <TR> 
                              <TD class=S align=right height=25>会员密码</TD>
                              <TD> 
                                <INPUT class=inputt01 id=password3 
                              type=password size=14 value=<%=Password%>
                            name=Pwd>
                              </TD>
                            </TR>
                            <TR> 
                              <TD></TD>
                              <TD>&nbsp;</TD>
                              <TD></TD>
                            </TR>
                          </FORM>
                        </TABLE>
                        <BR>
                      </TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                  <TABLE cellSpacing=0 cellPadding=0 border=0>
                    <TBODY> 
                    <TR> 
                      <TD width=323> 
                        <P>&nbsp;</P>
                      </TD>
                    </TR>
                    </TBODY> 
                  </TABLE>
                  <TABLE cellSpacing=1 borderColorDark=#ffffff cellPadding=0 
                  width=333 bgColor=#bababa border=0>
                    <TBODY> 
                    <TR> 
                      <TD width=331 bgColor=white> 
                        <TABLE cellSpacing=0 borderColorDark=#ffffff 
                        cellPadding=0 width="100%" bgColor=#f3f3f3 border=0>
                          <TBODY> 
                          <TR> 
                            <TD width=15 bgColor=#e53c00> 
                              <P>&nbsp;</P>
                            </TD>
                            <TD align=left height=30><FONT 
                              color=white>&nbsp;&nbsp;</FONT><SPAN 
                              class=font14>最新供应信息</SPAN></TD>
                            <TD width=100><IMG height=20 
                              src="images/post1.gif" 
                          width=95></TD>
                          </TR>
                          <TR> 
                            <TD background=images/line.gif 
                            colSpan=3 height=1><FONT 
                          face=宋体></FONT></TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                        <TABLE id=datalistbuy 
                        style="WIDTH: 100%; BORDER-COLLAPSE: collapse" 
                        cellSpacing=0 border=0>
                          <TBODY> 
                          <TR> 
                            <TD> 
                              <TABLE cellSpacing=0 cellPadding=1 width="99%" 
                              border=0>
                                <TBODY> 
                                <%
								set rs_info=conn.execute("select * from Info Where InfoType='供应' Order by InfoID desc")
								if rs_info.eof and rs_info.bof then
								%>
                                <TR> 
                                  <TD width=10 height=28><B><FONT color=#d50000 
                                size=1>&gt;</FONT></B></TD>
                                  <TD style="LINE-HEIGHT: 140%" align=left>No 
                                    Data!</TD>
                                </TR>
                                <%
								else
								i=0
							    do while not rs_info.eof
								%>
                                <TR> 
                                  <TD width=10 height=28><B><FONT color=#d50000 
                                size=1>&gt;</FONT></B></TD>
                                  <TD style="LINE-HEIGHT: 140%" align=left><A 
                                href="../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>"><%=rs_info("title")%></A>&nbsp;</TD>
                                </TR>
                                <%
								i=i+1
								if i>=MaxNum then exit do
								rs_info.movenext
								loop
								end if
								rs_info.Close
								set rs_info=nothing
								%>
                                </TBODY> 
                              </TABLE>
                            </TD>
                          </TR>
                          <TR> </TR>
                          </TBODY> 
                        </TABLE>
                      </TD>
                    </TR>
                    <TR> 
                      <TD width=331 bgColor=white> 
                        <TABLE cellSpacing=0 borderColorDark=#ffffff 
                        cellPadding=0 width="100%" bgColor=#f3f3f3 border=0>
                          <TBODY> 
                          <TR> 
                            <TD width=10 bgColor=#ff7316></TD>
                            <TD align=left height=30><FONT 
                              color=white>&nbsp;&nbsp;</FONT><SPAN 
                              class=font14>最新求购信息</SPAN></TD>
                            <TD width=100><IMG height=20 
                              src="images/post2.gif" 
                          width=95></TD>
                          </TR>
                          <TR> 
                            <TD background=images/line.gif 
                            colSpan=3 height=1></TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                        <TABLE id=datalistbuy 
                        style="WIDTH: 100%; BORDER-COLLAPSE: collapse" 
                        cellSpacing=0 border=0>
                          <TBODY> 
                          <TR> 
                            <TD> 
                              <TABLE cellSpacing=0 cellPadding=1 width="99%" 
                              border=0>
                                <TBODY> 
                                <%
								set rs_info=conn.execute("select * from Info Where InfoType='采购' Order by InfoID desc")
								if rs_info.eof and rs_info.bof then
								%>
                                <TR> 
                                  <TD width=10 height=28><B><FONT color=#d50000 
                                size=1>&gt;</FONT></B></TD>
                                  <TD style="LINE-HEIGHT: 140%" align=left>No 
                                    Data!&nbsp;</TD>
                                </TR>
                                <%
								else
								i=0
							    do while not rs_info.eof
								%>
                                <TR> 
                                  <TD width=10 height=28><B><FONT color=#d50000 
                                size=1>&gt;</FONT></B></TD>
                                  <TD style="LINE-HEIGHT: 140%" align=left><A 
                                href="../Trade/InfoView.asp?InfoID=<%=rs_info("InfoID")%>"><%=rs_info("title")%></A>&nbsp;</TD>
                                </TR>
                                <%
								i=i+1
								if i>=MaxNum then exit do
								rs_info.movenext
								loop
								end if
								rs_info.Close
								set rs_info=nothing
								%>
                                </TBODY> 
                              </TABLE>
                            </TD>
                          </TR>
                          <TR> </TR>
                          </TBODY> 
                        </TABLE>
                      </TD>
                    </TR>
                    <TR> 
                      <TD width=331 bgColor=white> 
                        <TABLE cellSpacing=0 borderColorDark=#ffffff 
                        cellPadding=0 width="100%" bgColor=#f3f3f3 border=0>
                          <TBODY> 
                          <TR> 
                            <TD width=5 bgColor=#ff9013> 
                              <P>&nbsp;</P>
                            </TD>
                            <TD align=left height=30><FONT 
                              color=white>&nbsp;&nbsp;</FONT><SPAN 
                              class=font14>最新加入公司库公司</SPAN></TD>
                            <TD width=100><IMG height=20 
                              src="images/post3.gif" 
                          width=95></TD>
                          </TR>
                          <TR> 
                            <TD background=images/line.gif 
                            colSpan=3 height=1></TD>
                          </TR>
                          </TBODY> 
                        </TABLE>
                        <TABLE id=dlistcom 
                        style="WIDTH: 100%; BORDER-COLLAPSE: collapse" 
                        cellSpacing=0 border=0>
                          <TBODY> 
                          <TR> 
                            <TD> 
                              <table cellspacing=0 cellpadding=1 width="99%" 
                              border=0>
                                <tbody> 
                                <%
								set rs_q=conn.execute("select * from qyml Where UFlag=1 Order by ID desc")
								if rs_q.eof and rs_q.bof then
								%>
                                <tr> 
                                  <td width=10 height=28><b><font color=#d50000 
                                size=1>&gt;</font></b></td>
                                  <td style="LINE-HEIGHT: 140%" align=left>No 
                                    Data!&nbsp;</td>
                                </tr>
                                <%
								else
								i=0
							    do while not rs_q.eof
								%>
                                <tr> 
                                  <td width=10 height=28><b><font color=#d50000 
                                size=1>&gt;</font></b></td>
                                  <td style="LINE-HEIGHT: 140%" align=left><a 
                                href="../ShowRoom/Index.asp?GsID=<%=rs_q("ID")%>"><%=rs_q("Qymc")%></a>&nbsp;</td>
                                </tr>
                                <%
								i=i+1
								if i>=MaxNum then exit do
								rs_q.movenext
								loop
								end if
								rs_q.Close
								set rs_q=nothing
								%>
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
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD>
  </TR>
  </TBODY> 
</TABLE>
<!--#include file="../inc/down.htm"-->
</BODY>
</HTML>
<%
end sub
          
conn.close
set conn=nothing
%>