<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<% 
if request.querystring="addsave" then

dim Server_V1,Server_V2
Server_V1=Cstr(Request.ServerVariables("HTTP_REFERER"))
Server_V2=Cstr(Request.ServerVariables("SERVER_NAME"))
If Mid(Server_V1,8,Len(Server_V2))<>Server_V2 Then Call ShowError("�Բ������ⲿ�ύ���ݣ�")

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
    Response.Write "<script>alert('�ύ���Գɹ���');this.location.href='Index.asp'</script>"
End if
%>
<HTML>
<HEAD>
<TITLE>������ѯ������������</TITLE>
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
	alert("���������������Ŷ��");
	lybook.U_Name.focus();
	return false;
	}
	
  	if (lybook.U_Tel.value==""){
	alert("�����������绰Ŷ��");
	lybook.U_Tel.focus();
	return false;
	}
	
  	if (lybook.U_Email.value==""){
	alert("�����ʼ�һ��Ҫ��дŶ��");
	lybook.Email.focus();
	return false;
	}
	
	haha=lybook.U_Email.value
    if(haha.length>0)
	{
        i=haha.indexOf("@");
        if(i==-1)
		{
			alert("�ۣ�������ĵ����ʼ��д���Ŷ��")
			lybook.U_Email.focus();
		    return false;
        }
       ii=haha.indexOf(".")
        if(ii==-1)
		{
			alert("�ۣ�������ĵ����ʼ��д���Ŷ��")
			lybook.U_Email.focus();
			return false;
        }
    }
	if (checktext(haha))
	{
		alert("�ܸ�����������Ч�����ʼ���");
		lybook.U_Email.focus();
		return false;
	}
	
  	if (lybook.U_Content.value==""){
	alert("���Ի�ûд�أ������ɲ���Ŷ����");
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
                  <TD align=right width=60>������</TD>
                  <TD> 
                    <INPUT class=button1 maxLength=15 name=U_Name>
                    <SPAN 
                  class=font1><SPAN class="font2 style2">��</SPAN></SPAN></TD>
                </TR>
                <TR> 
                  <TD align=right>�绰��</TD>
                  <TD> 
                    <INPUT class=button1 maxLength=50 name=U_Tel>
                    <SPAN 
                  class=font1><SPAN class="font2 style2">��</SPAN></SPAN></TD>
                </TR>
                <TR> 
                  <TD align=right>���䣺</TD>
                  <TD> 
                    <INPUT class=button1 maxLength=50 
                  name=U_Email>
                    <SPAN class=font1><span class="font2 style2">��</span></SPAN></TD>
                </TR>
                <TR> 
                  <TD align=right>���ԣ�</TD>
                  <TD> 
                    <select name=U_ComeFrom>
                      <option 
selected>==��ѡ��==</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=�Ϻ�>�Ϻ�</option>
                      <option value=���>���</option>
                      <option value=����>����</option>
                      <option value=�ӱ�>�ӱ�</option>
                      <option value=ɽ��>ɽ��</option>
                      <option value=���ɹ�>���ɹ�</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=������>������</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=ɽ��>ɽ��</option>
                      <option value=�Ĵ�>�Ĵ�</option>
                      <option value=�㽭>�㽭</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=�㶫>�㶫</option>
                      <option value=̨��>̨��</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=�ຣ>�ຣ</option>
                      <option value=����>����</option>
                      <option value=�½�>�½�</option>
                      <option value=���>���</option>
                      <option value=����>����</option>
                      <option value=̨��>̨��</option>
                      <option value=����>����</option>
                      <option value=�¼���>�¼���</option>
                      <option value=�¹�>�¹�</option>
                      <option value=����>����</option>
                      <option value=Ӣ��>Ӣ��</option>
                      <option value=���ô�>���ô�</option>
                      <option value=������>������</option>
                      <option 
value=�Ĵ�����>�Ĵ�����</option>
                      <option value=����>����</option>
                      <option value=����>����</option>
                      <option value=�ձ�>�ձ�</option>
                      <option value=�����>�����</option>
                      <option 
value=����˹>����˹</option>
                      <option 
value=����>����</option>
                    </select>
                    <SPAN 
                  class=font1></SPAN></TD>
                </TR>
                <TR> 
                  <TD align=right>���⣺</TD>
                  <TD> 
                    <INPUT class=button1 maxLength=25 name=U_Title>
                    <SPAN 
                  class=font1><SPAN class="font2 style2">��</SPAN></SPAN></TD>
                </TR>
                <TR> 
                  <TD vAlign=middle align=right>���ݣ�</TD>
                  <TD> 
                    <textarea class=button1 name=U_Content rows=6 wrap=PHYSICAL cols=28 0></textarea>
                  </TD>
                </TR>
                <TR> 
                  <TD align=right>&nbsp;</TD>
                  <TD> 
                    <INPUT class=button2 type=submit value=д���� name=Submit>
                    <INPUT class=button2 onClick=" history.back()" type=button value=��д�� name=cmdExit>
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
          <TD class=font1><FONT size=2>��</FONT>��<SPAN 
            class="font2 style2">��</SPAN>�ŵ�Ϊ������ </TD>
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
strUnit="����Ϣ"
sfilename="index.asp"
    sql="select * from Ys_Message order by YsID desc"
	Set Rs=server.createobject("Adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	    response.write "<TABLE cellSpacing=0 cellPadding=0 width=680 height=60 border=0><TR><TD align=center height=33><SPAN class=font1>������</SPAN></TD></TR></TABLE>"
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
                background=images/pin-t-name.jpg>����:</td>
                <td class=font2 width="115">&nbsp;<font color="#CC6600"><%=rs("YsName")%></font></td>
                <td class=font2 width="239">����:&nbsp;<font color="#CC6600"><%=rs("YsTitle")%></font></td>
                <td class=font2 width="240">����:&nbsp;<font color="#CC6600"><%=rs("YsComeFrom")%></font></td>
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
                background=images/pin-con-01.jpg>��������</td>
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
                      <td class=font2 align=right width=40>ʱ��:</td>
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
                <td class=font1>�ͷ��ظ�</td>
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