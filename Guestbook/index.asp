<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<% 
if request.querystring="addsave" then

dim Server_V1,Server_V2
Server_V1=Cstr(Request.ServerVariables("HTTP_REFERER"))
Server_V2=Cstr(Request.ServerVariables("SERVER_NAME"))
If Mid(Server_V1,8,Len(Server_V2))<>Server_V2 Then Call ShowError("对不起不能外部提交数据！")

    set rs=server.createobject("adodb.recordset")
    sqltext="select * from Ys_Message"
    rs.open sqltext,conn,1,3
    rs.addnew 
    rs("YsName")=server.HTMLEncode(Request("U_Name"))
    rs("YsTel")=server.HTMLEncode(Request("U_Tel"))
    rs("YsEmail")=server.HTMLEncode(Request("U_Email"))
    rs("YsComeFrom")=server.HTMLEncode(Request("U_ComeFrom"))
    rs("YsTitle")=server.HTMLEncode(Request("U_Title"))
    rs("YsContent")=dvHTMLEncode(Request("U_Content"))
    rs("YsAddTime")=cstr(now())
    rs("YsIP")=Request.ServerVariables("Remote_Addr")
    rs.update
    rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    Response.Write "<script>alert('提交留言成功！');this.location.href='Index.asp'</script>"
End if
%>
<HTML>
<HEAD>
<TITLE>在线咨询－诚信物流网</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="../images/mian.css" type=text/css rel=stylesheet>
<STYLE type=text/css>.right {
	BORDER-RIGHT: #ff7300 1px solid
}
.net {
	BORDER-RIGHT: #f79218 1px solid; BORDER-TOP: #f79218 1px solid; BORDER-LEFT: #f79218 1px solid; BORDER-BOTTOM: #f79218 1px solid
}
.mid {
	WIDTH: 2px
}
.font1 {
	PADDING-LEFT: 20px; FONT-SIZE: 12px; COLOR: #333333; LINE-HEIGHT: 21px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif
}
.font2 {
	FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif
}
.font3 {
	PADDING-RIGHT: 20px; PADDING-LEFT: 20px; FONT-SIZE: 12px; COLOR: #333333; LINE-HEIGHT: 25px; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif
}
.mid1 {
	BORDER-TOP: #cccccc 1px dashed; LINE-HEIGHT: 2px
}
</STYLE>

<SCRIPT language=javascript>
<!--
  function mOvr(src,clrOver) {
    if (!src.contains(event.fromElement)) {
	  src.bgColor = clrOver;
	}
  }
  function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
	  src.bgColor = clrIn;
	}
  }
  
function datacheck()
{  
  	if (lybook.U_Name.value==""){
	alert("你好像还忘了填姓名哦！");
	lybook.U_Name.focus();
	return false;
	}
	
  	if (lybook.U_Tel.value==""){
	alert("你好像还忘了填电话哦！");
	lybook.U_Tel.focus();
	return false;
	}
	
  	if (lybook.U_Email.value==""){
	alert("电子邮件一定要填写哦！");
	lybook.Email.focus();
	return false;
	}
	
	haha=lybook.U_Email.value
    if(haha.length>0)
	{
        i=haha.indexOf("@");
        if(i==-1)
		{
			alert("哇！您输入的电子邮件有错误哦！")
			lybook.U_Email.focus();
		    return false;
        }
       ii=haha.indexOf(".")
        if(ii==-1)
		{
			alert("哇！您输入的电子邮件有错误哦！")
			lybook.U_Email.focus();
			return false;
        }
    }
	if (checktext(haha))
	{
		alert("能告诉我您的有效电子邮件吗？");
		lybook.U_Email.focus();
		return false;
	}
	
  	if (lybook.U_Content.value==""){
	alert("留言还没写呢，这样可不行哦！！");
	lybook.U_Content.focus();
	return false;
	}

}
function checktext(text)
{
		allValid = true;
		for (i = 0;  i < text.length;  i++)
		{
			if (text.charAt(i) != " ")
			{
				allValid = false;
				break;
			}
		}
return allValid;
}
-->
</SCRIPT>
<LINK href="images/css.css" type=text/css rel=stylesheet>
<STYLE type=text/css>.style2 {
	COLOR: #ff0000
}
</STYLE>

<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff topMargin=0>
<!--#include file="top1.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY> 
  <TR> 
    <TD class=net vAlign=top width=286> 
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY> 
        <TR> 
          <TD><IMG height=47 src="images/pin-write.jpg" 
        width=286></TD>
        </TR>
        <TR> 
          <TD> 
            <FORM name=lybook onSubmit="return datacheck();" action="index.asp?addsave" method=post>
              <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TBODY> 
                <TR> 
                  <TD align=right width=60>姓名：</TD>
                  <TD> 
                    <INPUT class=button1 maxLength=15 name=U_Name>
                    <SPAN 
                  class=font1><SPAN class="font2 style2">√</SPAN></SPAN></TD>
                </TR>
                <TR> 
                  <TD align=right>电话：</TD>
                  <TD> 
                    <INPUT class=button1 maxLength=50 name=U_Tel>
                    <SPAN 
                  class=font1><SPAN class="font2 style2">√</SPAN></SPAN></TD>
                </TR>
                <TR> 
                  <TD align=right>邮箱：</TD>
                  <TD> 
                    <INPUT class=button1 maxLength=50 
                  name=U_Email>
                    <SPAN class=font1><span class="font2 style2">√</span></SPAN></TD>
                </TR>
                <TR> 
                  <TD align=right>来自：</TD>
                  <TD> 
                    <select name=U_ComeFrom>
                      <option 
selected>==请选择==</option>
                      <option value=河南>河南</option>
                      <option value=北京>北京</option>
                      <option value=上海>上海</option>
                      <option value=天津>天津</option>
                      <option value=重庆>重庆</option>
                      <option value=河北>河北</option>
                      <option value=山西>山西</option>
                      <option value=内蒙古>内蒙古</option>
                      <option value=辽宁>辽宁</option>
                      <option value=吉林>吉林</option>
                      <option value=黑龙江>黑龙江</option>
                      <option value=江苏>江苏</option>
                      <option value=安徽>安徽</option>
                      <option value=福建>福建</option>
                      <option value=江西>江西</option>
                      <option value=山东>山东</option>
                      <option value=四川>四川</option>
                      <option value=浙江>浙江</option>
                      <option value=湖北>湖北</option>
                      <option value=云南>云南</option>
                      <option value=湖南>湖南</option>
                      <option value=广东>广东</option>
                      <option value=台湾>台湾</option>
                      <option value=广西>广西</option>
                      <option value=海南>海南</option>
                      <option value=贵州>贵州</option>
                      <option value=西藏>西藏</option>
                      <option value=陕西>陕西</option>
                      <option value=甘肃>甘肃</option>
                      <option value=青海>青海</option>
                      <option value=宁夏>宁夏</option>
                      <option value=新疆>新疆</option>
                      <option value=香港>香港</option>
                      <option value=澳门>澳门</option>
                      <option value=台湾>台湾</option>
                      <option value=美国>美国</option>
                      <option value=新加坡>新加坡</option>
                      <option value=德国>德国</option>
                      <option value=法国>法国</option>
                      <option value=英国>英国</option>
                      <option value=加拿大>加拿大</option>
                      <option value=新西兰>新西兰</option>
                      <option 
value=澳大利亚>澳大利亚</option>
                      <option value=巴西>巴西</option>
                      <option value=韩国>韩国</option>
                      <option value=日本>日本</option>
                      <option value=意大利>意大利</option>
                      <option 
value=俄罗斯>俄罗斯</option>
                      <option 
value=其它>其它</option>
                    </select>
                    <SPAN 
                  class=font1></SPAN></TD>
                </TR>
                <TR> 
                  <TD align=right>主题：</TD>
                  <TD> 
                    <INPUT class=button1 maxLength=25 name=U_Title>
                    <SPAN 
                  class=font1><SPAN class="font2 style2">√</SPAN></SPAN></TD>
                </TR>
                <TR> 
                  <TD vAlign=middle align=right>内容：</TD>
                  <TD> 
                    <textarea class=button1 name=U_Content rows=6 wrap=PHYSICAL cols=28 0></textarea>
                  </TD>
                </TR>
                <TR> 
                  <TD align=right>&nbsp;</TD>
                  <TD> 
                    <INPUT class=button2 type=submit value=写好了 name=Submit>
                    <INPUT class=button2 onClick=" history.back()" type=button value=不写了 name=cmdExit>
                  </TD>
                </TR>
                </TBODY> 
              </TABLE>
            </FORM>
          </TD>
        </TR>
        <TR> 
          <TD><IMG height=48 src="images/pin-toknow.jpg" 
          width=285></TD>
        </TR>
        <TR> 
          <TD class=font1><FONT size=2>·</FONT>打<SPAN 
            class="font2 style2">√</SPAN>号的为必填项 </TD>
        </TR>
        <TR> 
          <TD>&nbsp;</TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD>
    <TD class=mid>&nbsp;</TD>
    <TD vAlign=top height="400"> 
      <TABLE cellSpacing=0 cellPadding=0 width=680 border=0>
        <TBODY> 
        <TR> 
          <TD align=right background=images/pin-list.jpg 
            height=33><SPAN class=font1></SPAN>&nbsp;&nbsp;</TD>
        </TR>
        </TBODY> 
      </TABLE>
      <%
dim rssum,maxpage,thepages,viewpage,i,strUnit,sfilename
maxpage=10	
strUnit="条信息"
sfilename="index.asp"
    sql="select * from Ys_Message order by YsID desc"
	Set Rs=server.createobject("Adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	    response.write "<TABLE cellSpacing=0 cellPadding=0 width=680 height=60 border=0><TR><TD align=center height=33><SPAN class=font1>无内容</SPAN></TD></TR></TABLE>"
	else
	    rssum=rs.recordcount
        if rssum mod maxpage > 0 then
            thepages=rssum\maxpage+1
        else
            thepages=rssum\maxpage
        end if

        viewpage=trim(request("page"))
        if viewpage="" or isnull(viewpage) or not(isnumeric(viewpage)) then
            viewpage=1
        elseif cint(viewpage)<1 then
            viewpage=1
        else
            if cint(viewpage)>thepages then
                viewpage=1
            end if
       end if
       i=1
       if viewpage<>"1" then
           rs.move (viewpage-1)*maxpage
       end if
       for i=1 to maxpage
           if rs.eof then exit for
 %>
      <table cellspacing=0 cellpadding=0 width=680 border=0>
        <tbody> 
        <tr> 
          <td> 
            <hr size=1>
          </td>
        </tr>
        <tr> 
          <td> 
            <table cellspacing=0 cellpadding=0 width=680 
            background=images/pin-t-bg.jpg border=0>
              <tbody> 
              <tr> 
                <td class=font2 align=right width=75 
                background=images/pin-t-name.jpg>姓名:</td>
                <td class=font2 width="115">&nbsp;<font color="#CC6600"><%=rs("YsName")%></font></td>
                <td class=font2 width="239">主题:&nbsp;<font color="#CC6600"><%=rs("YsTitle")%></font></td>
                <td class=font2 width="240">来自:&nbsp;<font color="#CC6600"><%=rs("YsComeFrom")%></font></td>
                <td width=11><img height=25 
                  src="images/pin-t-last.jpg" 
            width=11></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        <tr> 
          <td> 
            <table cellspacing=0 cellpadding=0 width=680 border=0>
              <tbody> 
              <tr> 
                <td class=font2 align=middle width=183 
                background=images/pin-con-01.jpg>留言内容</td>
                <td align=right 
                  background=images/pin-con-bg.jpg>&nbsp;</td>
                <td width=11><img height=21 
                  src="images/pin-con-last.jpg" 
            width=11></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        <tr> 
          <td> 
            <table cellspacing=0 cellpadding=0 width=680 
            background=images/pin-c-b-bg.jpg border=0>
              <tbody> 
              <tr> 
                <td class=font3 style="word-wrap: break-word;word-break: break-all"><%=rs("YsContent")%></td>
              </tr>
              <tr> 
                <td> 
                  <table cellspacing=0 cellpadding=0 width=680 border=0>
                    <tbody> 
                    <tr> 
                      <td>&nbsp;</td>
                      <td class=font2 align=right width=40>时间:</td>
                      <td class=font2 width=150><font color="#CC6600"><%=rs("YsAddTime")%></font></td>
                      <td width=14>&nbsp;</td>
                      <td 
            width=10>&nbsp;</td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        <%if rs("YsReply")<>"" then%>
        <tr> 
          <td align=middle background=images/pin-c-b-bg.jpg> 
            <table cellspacing=0 cellpadding=0 width="95%" border=0>
              <tbody> 
              <tr> 
                <td class=mid1>&nbsp;</td>
              </tr>
              <tr> 
                <td class=font1>客服回复</td>
              </tr>
              <tr> 
                <td class=font1><font 
                  color=red><%=rs("YsReply")%></font></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        <%end if%>
        <tr> 
          <td><img height=12 src="images/pin-c-l.jpg" 
        width=680></td>
        </tr>
        </tbody> 
      </table>
      <%
           rs.movenext
       next
%>
      <TABLE cellSpacing=0 cellPadding=0 width=680 
      background=images/pin-down.jpg border=0>
        <TBODY> 
        <TR> 
          <TD class=font1 align=middle height=33> 
            <%call showpage (strUnit,sfilename,10,"#ff0000")%>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD>
  </TR>
  <%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  </TBODY> 
</TABLE>
<!--#include file="bottom.htm"-->
</BODY></HTML>