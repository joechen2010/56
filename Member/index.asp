<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->

<%logintimes=session("logintimes")
'set rs=server.createobject("adodb.recordset")
'sql="select * from qyml where User='"&session("user")&"'"
'rs.open sql,conn,1,3
'if rs.eof and rs.bof then
'response.write"<SCRIPT language=JavaScript>alert('�Բ��𣬴���Ϣ�����ڣ�');"
'response.write"javascript:history.go(-1)<SCRIPT>"
'response.end 
'else
'rs("LastLoginIP")=Request.ServerVariables("REMOTE_ADDR")
'rs("LastLoginTime")=now()
'rs("LoginTimes")=rs("LoginTimes")+1
'rs.update
'rs.close
'set rs=nothing
'end if
'conn.execute("update qyml set LastLoginIP='"&Request.ServerVariables("REMOTE_ADDR")&"',LastLoginTime='"&now&"',LoginTimes=LoginTimes+1 Where User='"&UsernameGet&"'")
%>

<%
   set rs_m=conn.execute("select count(*) from message where New=0 and T_User='"&session("user")&"'")
   set rs_m2=conn.execute("select count(*) from message where T_User='"&session("user")&"'")
   set rs_p=conn.execute("select count(*) from info2 where gsid='"&session("user")&"'")
   set rs_p1=conn.execute("select count(*) from spzs where flag=1 and gsid='"&session("user")&"'")
   set rs_p2=conn.execute("select count(*) from spzs where flag=0 and gsid='"&session("user")&"'")
		   
   set rs_info=conn.execute("select count(*) from info where gsid='"&session("user")&"'")
   set rs_info1=conn.execute("select count(*) from info where flag=1 and gsid='"&session("user")&"'")
   set rs_info2=conn.execute("select count(*) from info where flag=0 and gsid='"&session("user")&"'")
   
   set rs_h=conn.execute("select count(*) from Qyzs where gsid='"&session("user")&"'")
   set rs_car=conn.execute("select count(*) from info2 where infotype='��Դ��Ϣ' and gsid='"&session("user")&"'")
   set goods=conn.execute("select count(*) from info2 where infotype='��Դ��Ϣ' and gsid='"&session("user")&"'")
   set zx_k=conn.execute("select count(*) from zx_info where infotype='����ר��' and gsid='"&session("user")&"'")
   set zx_h=conn.execute("select count(*) from zx_info where infotype='����ר��' and gsid='"&session("user")&"'")
   set zp=conn.execute("select count(*) from zp_info where gsid='"&session("user")&"'")
   set repair=conn.execute("select count(*) from repair_info where gsid='"&session("user")&"'")
   set banyun=conn.execute("select count(*) from banyun_info where gsid='"&session("user")&"'")
   set files=conn.execute("select count(*) from file_info where gsid='"&session("user")&"'")
   set gq=conn.execute("select count(*) from gq_info where gsid='"&session("user")&"'")
   set kf=conn.execute("select count(*) from kf_info where gsid='"&session("user")&"'")
%>
<HTML>
<HEAD>
<TITLE>���������� ��Ա�������� - �������˵����ϼ�԰</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>
BODY {
	MARGIN: 0px; BACKGROUND-COLOR: #2c68b1
}
.style1 {
	COLOR: #000000
}
.style2 {
	COLOR: #ffffff
}
</STYLE>
<LINK href="images/main.css" type=text/css rel=stylesheet>
<SCRIPT language=JavaScript type=text/JavaScript>
<!--
function openwindow(url,width,height) { 	left1 =0; 	top1 = 0; 	window.open(url,"","width=" + width + ",height=" + height + ",toolbar=no,menubar=no,scrollbars=yes,location=no,status=yes,left=" + left1.toString() + ",top=" + top1.toString()); }
//resizable=yes, ���
//-->
</SCRIPT>
<META content="MSHTML 6.00.2462.0" name=GENERATOR></HEAD>
<BODY>
<!--#include file="head.htm"-->
<TABLE cellSpacing=0 cellPadding=0 width=778 border=0 align="center">
  <TBODY> 
  <TR>
    <TD vAlign=top width=160 bgColor=#2c68b1 height=566>
<!--#include file="left.htm"-->
      </TD>
    <TD id=main vAlign=top align=middle bgColor=#ffffff>
      <TABLE cellSpacing=0 cellPadding=3 width="94%" border=0>
        <TBODY>
        <TR>
          <TD id=pos align=left width="83%" height=38><IMG height=11 
            src="images/icon03.gif" width=15> <A 
            href="http://www.cx56w.com/">����������</A> &gt; ��Ա����������ҳ</TD>
          <TD width="17%"><IMG height=32 
            src="images/icon_onlineservice.gif" 
      width=132></TD>
        </TR></TBODY></TABLE>
      <table width="94%"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> </font></strong><b><span style="font-size: 10.5pt">�𾴵� 
            </span><font face=Arial color="#FF6600" style="font-size: 10.5pt"><%=session("name")%></font><span style="font-size: 10.5pt">�����ã� 
            <%if session("uflag")=0 then
		       response.write "����ʱ��û��ͨ�����!"
		  else
		        response.write "�������ǵ�" & ss &"!"
		  end if%>
            </span></b></td>
        </tr>
      </table>
      <div align="center"> 
        <table cellspacing=0 cellpadding=0 width="90%" border=0>
          <tbody> 
          <tr> 
            <td width=10><img height=26 src="images/title_blue_1.gif" 
            width=10></td>
            <td class=titletx align=left width=537 
          background=images/title_blue_2.gif>������</td>
            <td width=8><img height=26 src="images/title_blue_3.gif" 
            width=8></td>
          </tr>
          </tbody> 
        </table>
        <table cellspacing=0 cellpadding=5 width="90%" border=0>
          <tbody> 
          <tr> 
            <td valign=top align=left height=90> 
              <div>��վ����<br>
                <table style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" cellspacing=0 cellpadding=1 width=100% border=0>
                  <%set rs_n=conn.execute("select top 5 * from NewsData where NClassID=7 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
                  <tr> 
                    <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No 
                      Data!</td>
                  </tr>
                  <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),50)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
                  <tr> 
                    <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="javascript:openwindow('../News/NewsDetail1.asp?NewsID=<%=rs_n("NewsID")%>',500,400)"><%=topic%></a> 
                      (<%=month(rs_n("RegTime"))%>/<%=day(rs_n("RegTime"))%>)</td>
                  </tr>
                  <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
                </table>			  
			  </div>
              <div align=right><a 
            href="#">���๫��...</a></div>
            </td>
          </tr>
          </tbody> 
        </table>
        <table cellspacing=0 cellpadding=0 width="90%" border=0>
          <tbody> 
          <tr> 
            <td width=10><img height=26 src="images/title_blue_1.gif" 
            width=10></td>
            <td class=titletx align=left width=537 
          background=images/title_blue_2.gif>ϵͳ״̬</td>
            <td width=8><img height=26 src="images/title_blue_3.gif" 
            width=8></td>
          </tr>
          </tbody> 
        </table>
        <table cellspacing=0 cellpadding=5 width="90%" border=0>
          <tbody> 
          <tr> 
            <td valign=top align=left height=90> 
              <div class=newsline>���� <img height=13 
            src="images/icon_msg.gif" width=18 align=absMiddle> <a href="inbox.asp"><%=rs_m(0)%></a> 
                ���������ԣ����� <a class="12linkBlue" href="inbox.asp"><%=rs_m2(0)%></a>�����ԣ�����鿴�������ԡ�</div>
              <div class=newsline>���� <a href="car_manage.asp"><%=rs_car(0)%></a> ����Դ��Ϣ������鿴���³�Դ��Ϣ��</div>
              <div class=newsline>���� <a href="goods_manage.asp"><%=goods(0)%></a> ����Դ��Ϣ������鿴���»�Դ��Ϣ��</div>
              <div class=newsline>���� <a href="kyzx_Manage.asp"><%=zx_k(0)%></a> ������ר�ߣ�����鿴���¿���ר�ߡ�</div>
              <div class=newsline>���� <a href="hyzx_Manage.asp"><%=zx_h(0)%></a> ������ר�ߣ�����鿴���»���ר�ߡ�</div>
              <div class=newsline>���� <a href="zp_manage.asp"><%=zp(0)%></a> ����Ƹ��Ϣ������鿴������Ƹ��Ϣ��</div>
              <div class=newsline>���� <a href="repair_manage.asp"><%=repair(0)%></a> ���������������鿴�������������</div>
              <div class=newsline>���� <a href="files_manage.asp"><%=files(0)%></a> ����������������鿴���³���������</div>
              <div class=newsline>���� <a href="gq_manage.asp"><%=gq(0)%></a> ��������Ϣ������鿴���¹�����Ϣ��</div>
              <div class=newsline>���� <a href="kf_manage.asp"><%=kf(0)%></a> ���ⷿ��Ϣ������鿴���¿ⷿ��Ϣ��</div>
              <div class=newsline>���� <a href="banyun_manage.asp"><%=banyun(0)%></a> �����˵�װ������鿴���°��˵�װ��</div>
                <br>
              <div class=newsline>���ǵ�<%=logintimes%>�ε�¼�����һ�ε�¼����<%=session("logintime")%>����ӭ��������</div>
            </td>
          </tr>
          </tbody> 
        </table>
      </div>
    </TD></TR></TABLE>
<!--#include file="bottom.htm"--></BODY></HTML>
