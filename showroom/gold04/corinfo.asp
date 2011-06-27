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
<TITLE><%=rs_q("qymc")%> - 物流通金牌企业 - 资信档案</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1106" name=GENERATOR></HEAD>
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
                        color=#ffffff><b>资信档案</b></font></a></td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="96%" border="0" cellspacing="1" cellpadding="0" bgcolor="f37600">
        <tr>
          <td bgcolor="f37600" height="3"></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">
            <table class=style height=40 cellspacing=0 cellpadding=0 width="100%" 
            align=center 
            border=0>
              <tr>
          <td height="20"></td>
        </tr>
        <tbody> 
        <tr> 
          <td bgcolor=#FFCC00 height="30"><span class=zi1>　会员活动记录 会员服务第 1 年</span></td>
        </tr>
        </tbody> 
      </table>
            <table class=style cellspacing=1 cellpadding=6 width="100%" 
            align=center border=0>
              <tr bgcolor=#f5f5f5> 
          <td width="24%" bgcolor=#fff4e5 align="left">已在诚信物流网购买的服务</td>
          <td width="61%" bgcolor=#F7F7F7> <font color="#FF0000"> 
            <%
				   if rs_q("flag")=0 then
	  ss="免费会员"
   elseif rs_q("flag")=1 then
	  ss="铜牌会员"
   elseif rs_q("flag")=2 then
	  ss="银牌会员"
   elseif rs_q("flag")=3 then
	  ss="金牌会员"
   end if
   response.write ss
				%>
            </font> (<%=rs_q("BeginDate")%> ~ <%=rs_q("enddate")%>)</td>
        </tr>
        <tbody> 
        <tr bgcolor=#f5f5f5> 
          <td width="24%" bgcolor=#fff4e5 align="left"> 注册会员的时间</td>
          <td width="61%" bgcolor=#F7F7F7><%=rs_q("idate")%></td>
        </tr>
        <tr bgcolor=#f5f5f5> 
          <td bgcolor=#fff4e5> 最近登录的时间</td>
          <td bgcolor=#F7F7F7><%=rs_q("LastLoginTime")%></td>
        </tr>
        <tr bgcolor=#f5f5f5> 
          <td bgcolor=#fff4e5 align="left" rowspan="2"> 已发布上网的信息数量</td>
          <td bgcolor=#F7F7F7>供求商机：<font color=#ff0000>&nbsp; </font><font color=#ff0000><%=conn.execute("select count(*) from info where gsid='"&rs_q("user")&"'")(0)%></font></td>
        </tr>
        <tr bgcolor=#f5f5f5> 
          <td bgcolor=#F7F7F7>展示产品：<font color=#ff0000> 
            <%set rs_p=conn.execute("select count(*) from spzs where gsid='"&rs_q("user")&"'")
				  response.write rs_p(0)
				  rs_p.close
				  set rs_p=nothing%>
            </font></td>
        </tr>
        </tbody> 
      </table>
            <table class=style height=30 cellspacing=0 cellpadding=0 width="100%" 
            align=center bgcolor=#fafafa border=0>
              <tbody> 
              <tr> 
          <td bgcolor=#FFCC00><span 
            class=zi1>　证书及荣誉</span></td>
        </tr>
        </tbody> 
      </table>
            <table class=style cellspacing=2 cellpadding=2 width="100%" 
            align=center border=0>
              <tbody> 
              <tr bgcolor=#f5f5f5> 
              <td width="49%" height="89" align=middle> 
                <%set QyzsRs=Conn.Execute("Select * from qyzs where gsid='"&rs_q("user")&"'")
				   if QyzsRs.eof and QyzsRs.bof then
				       response.write "该公司还未发布企业证书!"
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
            <table class=style cellspacing=0 cellpadding=4 width="100%" 
            align=center border=0>
              <tbody> 
              <tr> 
          <td 
                  bgcolor=#fff4e5>请注意：<br>
            诚信物流网仅列示上述信息，上述报告仅代表报告出具日的情况，并不保证您使用上述信息时的情况与之相符，也不担保该信息的准确性，完整性和及时性，也不承担浏览者的任何商业风险。因此，诚信物流网不承担您因此而发生或致使的任何损害。诚信物流网保留全部或部分删除上述报告的权利。上述页面内容禁止以各种形式被全部或部分复制使用。 
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
        </tr>
      </table>
      <h3>&nbsp;</h3>
    </td>
  </tr>
</table>
        


<!--#include file="Copyright.htm"-->
</BODY></HTML>
<%end if%>
