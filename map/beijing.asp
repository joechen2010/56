<!--#include file="../Inc/Conn.asp"-->
<!--#include file="../Inc/Function.asp"-->
<%
const MaxNumber=50
const MaxNum=8
const MaxLen=28
dim rs_sel,sel1_total,rs_buy,buy_total,rs_p,p_total
set rs_sel=conn.execute("select Count(*) from Info where InfoType='��Ӧ'")
sell_total=rs_sel(0)
rs_sel.close
set rs_sel=nothing

set rs_buy=conn.execute("select Count(*) from Info where InfoType='�ɹ�'")
buy_total=rs_buy(0)
rs_buy.close
set rs_buy=nothing

set rs_s_o=conn.execute("select Count(*) from offer where OfferType='��Ӧ'")
sell_o_total=rs_s_o(0)
rs_s_o.close
set rs_s_o=nothing

set rs_b_o=conn.execute("select Count(*) from offer where OfferType='�ɹ�'")
buy_o_total=rs_b_o(0)
rs_b_o.close
set rs_b_o=nothing

set rs_p=conn.execute("select Count(*) from qyml")
p_total=rs_p(0)
rs_p.close
set rs_p=nothing
%>
<HTML>
<HEAD>
<TITLE>168������Ϣ��--ȫ����רҵ��������Ϣ����ƽ̨</TITLE>
<LINK href="../images/css.css" type=text/css rel=stylesheet>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content=���š����ų��ų�����������/����������Ϣ/����/��Դ��Ϣ/��Դ��Ϣ�����/�����,����/������/������Ϣ��,����/�������������Ż�����Ϣ/����������Ϣ,����������Ϣ��/������Ϣ��,������Ϣ����ƽ̨/����������Ϣ,����/���. 
name=description>
<META  content=���ţ���������������������Ϣ������������Ϣ�������ţ�cx56w,������Ϣ������������������������Ϣ�����������Ϣ����������/����������Ϣ/����/��Դ��Ϣ/��Դ��Ϣ�����/�����,����/������/������Ϣ��,����/�������������Ż�����Ϣ/����������Ϣ,����������Ϣ��/������Ϣ��,������Ϣ����ƽ̨/����������Ϣ,����/���,������������Ϣ��������Ϣ���ף�������Ϣ����ƽ̨������������Ϣ������������������Ϣ����ƽ̨��ȫ��������ȫ��������Ϣ���й��������й�������Ϣ����������������Ϣ����ȫ����������ȫ��������Ϣ�����й����������й�������Ϣ��������������Ϣ��ȫ�������ȫ�������Ϣ���й������ 
name=keywords �й�������������Ϣ�����������������Ϣ��������վ���֣���Դ��Ϣ����Դ��Ϣ���й����������� 
���¼���������������,���߷����Ϻ�����,��ҵ����,��ݸ����,��������,�ɶ�����,�㶫����,�㽭����,��������,��������,��������,�л�������,�й���������,�й���ҵ����,�绰�Ų�,��������,��������,�Ų�,��ͷ����,��������,114,��������,��������,��������,��������,��������,�ƶ�����,������ѯ,��������,����114,��ҵ����,�����ƹ�,����Ӫ��,��С��ҵ,��ҵĿ¼,��������,ȫ������,��������,����ָ��,��ѧ,����,�Ļ�,ҽҩ,����,����,������ҵ,�������,����,����,����,����,�Ӽ�������Ʒ,��װ,Ƥ��,��֯,����������Ʒ,��Ʒ,����Ʒ,�������,���,����,ӡˢ,��װ,��ֽ,����,����,���ز�,װ��,�����,������,ͨѶ�칫�豸����Ʒ,����,�繤,����,��е,�豸,����,�Ǳ�רҵ��Ʒ,����,�ǽ������ϼ���Ʒ,����,��ͨ,��Դ,���,ұ��ұ��,����,ʳƷ,ũ������ 
��������������Ϣ����ȫ����������ȫ��������Ϣ�����й����������й�������Ϣ�������ţ���������������������Ϣ������������������������Ϣ�����й����������Ϣ�������ģ�������Ϣ�ͻ��ˣ�������Ϣ���ģ��й��������������������ˣ�����������,����,�й�����,����,�й�����,�Ϻ�����,��ҵ����,��ݸ����,��������,�ɶ�����,�㶫����,�㽭����,��������,��������,��������,�л�������,�й���������,�й���ҵ����,�绰�Ų�,��������,��������,�Ų�,��ͷ����,��������,114,��������,��������,��������,��������,��������,�ƶ�����,ӡˢ����,������ѯ,��������,����114,��ҵ����,�����ƹ�,����Ӫ��,��С��ҵ,��ҵĿ¼,��������,ȫ������,��������,����ָ��,��ѧ,����,�Ļ�,ҽҩ,����,����,������ҵ,�������,����,����,����,����,�Ӽ�������Ʒ,��װ,Ƥ��,��֯,����������Ʒ,��Ʒ,����Ʒ,�������,���,����,ӡˢ,��װ,��ֽ,����,����,���ز�,װ��,�����,������,ͨѶ�칫�豸����Ʒ,����,�繤,����,��е,�豸,����,�Ǳ�רҵ��Ʒ,��������,��ҵ����,������Ϣ,��ҵ�����,��ҵó�׻���,������ó�׻���,ó�����飬�������,�������,�������,��Ϣ��ѯ,����ע��,�ռ�����,��վ����,��������,�й���?������Ϣ,���ñ���,��ҵ����,ͳ������,��ҵ����,ÿ���̻�,ó�׹�����,��ҵ��������,��ҵ�Ż�,��ҵ��վ�߻�����վ��ҳ������ά���������������������ü����������� 
������������Ϣ����ȫ���������ȫ�������Ϣ�����й���������й������Ϣ�������ˣ�������Ϣ��ȫ�����ˣ�ȫ��������Ϣ���й����ˣ��й�������Ϣ��>
<SCRIPT language=JavaScript type=text/JavaScript>
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</SCRIPT>
<META content="MSHTML 6.00.2462.0" name=GENERATOR>

</HEAD>
<BODY topMargin=0>
<TABLE cellSpacing=0 cellPadding=0 width=778 align=center border=0 bgcolor="#FFFFFF">
  <TBODY> 
  <TR> 
    <TD vAlign=center align=middle width=203> 
      <div align="center"><a 
      href="http://www.cx56w.com"><img src="../images/logo.gif" border=0><br>
        </a><font color="#FF0000">�����ǣ� 
        <script language=javascript>                                     
var day="";                                      
var month="";                                      
var ampm="";                                      
var ampmhour="";                                      
var myweekday="";                                      
var year="";                                      
mydate = new Date();                                      
myweekday = mydate.getDay();                                      
mymonth = mydate.getMonth();                                      
myday= mydate.getDate();                                      
myyear= mydate.getYear();                                      
myhours = mydate.getHours();                                      
ampmhour  =  (myhours > 12) ? myhours - 12 : myhours;                                      
if (ampmhour == "0") ampmhour = 0;                                      
ampm =  (myhours >= 12) ? ' ����' : ' ����';                                      
mytime = mydate.getMinutes();                                      
myminutes =  ((mytime < 10) ? ':0' : ':') + mytime;                                      
year = (myyear > 200) ? myyear : 1901 + myyear;                                      
if(myweekday == 0)                                      
weekday = " ������ ";                                      
else if(myweekday == 1)                                      
weekday = " ����һ ";                                      
else if(myweekday == 2)                                      
weekday = " ���ڶ� ";                                      
else if(myweekday == 3)                                      
weekday = " ������ ";                                      
else if(myweekday == 4)                                      
weekday = " ������ ";                                      
else if(myweekday == 5)                                      
weekday = " ������ ";                                      
else if(myweekday == 6)       weekday = " ������ ";                                      
                            </script>
        <script>                                        
document.write(year + "��" + (mymonth+1) + "��" +myday + "��")            
document.write(" " +weekday + " ");</script>
        </font></div>
    </TD>
    <TD vAlign=top width=8></TD>
    <TD vAlign=top align=right>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td valign=center width="13%"> 
            <div align="right"><span 
            style="FONT-SIZE: 12px"></span><a 
            href="../Reg/User_Reg.asp" target=_blank><img 
            height=17 src="../images/register_button.gif" width=60 
            border=0></a></div>
          </td>
          <td valign=center align=middle width="12%"> 
            <div align="center"><span 
            style="FONT-SIZE: 12px"><font color=#000000><a 
            href="../Login/login.asp" target=_blank><img 
            height=17 src="../images/login_button.gif" width=40 
            border=0></a></font></span></div>
          </td>
          <td valign=center class=s height="30" width="75%" align="right"> 
            <div align="center"><a href="../index.asp">��ҳ</a> �� <a href="../aboutus/index.asp">��������</a> 
              �� <a href="../agent/index.asp">���˷�վ</a> �� <a href="/union/index.asp">���ط�վ</a> 
              �� <a onMouseOver="javascript:window.external.addFavorite('http://www.cx56w.com','����������')" href="http://www.zhwlw.org/">�ղر�վ</a> 
              �� <a onMouseOver="javascript:this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.cx56w.com');"
href="http://www.cx56w.com" target="_self">��Ϊ��ҳ</a></div>

          </td>
        </tr>
      </table>
      <table cellspacing=0 cellpadding=0 border=0 width="100%">
        <tbody> 
        <tr align=right> 
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=92% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid" 
                align=center bgcolor=#f8f8f8><a 
                  href="../truck/index_beijing.asp" class=14link><strong><b>�һ��ҳ�</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=92% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid" 
                align=center bgcolor=#f8f8f8><a 
                  href="../Member/goods_add.asp" class=14link ><strong><b>������Դ</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid" 
                align=center bgcolor=#f8f8f8><a class=14link 
            href="file:///D|/Member/car_add.asp"><strong><b>������Դ</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid"><a class=14link 
            href="../Repairs/index_beijing.asp"><strong><b>�������</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid"><a class=14link 
            href="../job/index_beijing.asp"><strong><b>��Ƹ��ְ</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td bgcolor=#ff6600 height=6></td>
              </tr>
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid"><a href="../yellowpage/index_beijing.asp"  class=14link><b>������ҳ</b></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=30 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td height=6 bgcolor=#ff6600></td>
              </tr>
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #FF6600 1px solid; BORDER-LEFT: #FF6600 1px solid; BORDER-BOTTOM: #FF6600 1px solid"><a href="../news/index.asp" class=14link><b>������Ѷ</b></a><a 
                  href="/info/"></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
      <table cellspacing=0 cellpadding=0 border=0 width="100%">
        <tbody> 
        <tr align=right> 
          <td height="30"> 
            <table height=24 cellspacing=0 cellpadding=0 width=92% border=0>
              <tbody> 
              <tr> 
                <td 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
                align=center bgcolor=#f8f8f8><a 
                  href="../fline/index_beijing.asp" class=14link><strong><b>����ר��</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=92% border=0>
              <tbody> 
              <tr> 
                <td 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"
                align=center bgcolor=#f8f8f8><a 
                  href="../pline/index_beijing.asp" class=14link ><strong><b>����ר��</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"
                align=center bgcolor=#f8f8f8><a class=14link 
            href="../cheliang/index_beijing.asp"><strong><b>��������</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
               style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"><a class=14link 
            href="../info/index_beijing.asp"><strong><b>�ⷿ��Ϣ</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"><a class=14link 
            href="../trade/index_beijing.asp"><strong><b>������Ϣ</b></strong></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"><a href="index.asp"  class=14link><b>���ϵ�ͼ</b></a></td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td> 
            <table height=24 cellspacing=0 cellpadding=0 width=95% border=0>
              <tbody> 
              <tr> 
                <td 
                align=center bgcolor=#f8f8f8 
                style="BORDER-RIGHT: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid"><a href="../member" class=14link><b>��Ա����</b></a><a 
                  href="/info/"></a></td>
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
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
      <td align="center"><iframe id="ifrsearch" src="../search/index_car.asp" width="100%" frameborder="0" scrolling="no" height="90"></iframe></td>
    </tr>
  </table>
<table cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td height=5></td>
  </tr>
  </tbody> 
</table>
<TABLE cellSpacing=0 width=778 align=center border=0>
  <TBODY>
  <TR>
      <TD vAlign=top width=290> 
        <table width="275" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td height="180" valign="top"> 
              <TABLE cellSpacing=1 cellPadding=0 width="100%" bgColor=#c4c2c2 
        border=0>
		        <TR>
          <TD vAlign=top align=middle bgColor=#ffffff>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD background=../images/login_bg.jpg height=29>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD class=00088 width="45%" height=21>
                        <DIV class=style10 align=center><font color="#FFFFFF">����Ա��¼</font></DIV></TD>
                      <TD width="55%"></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=4 width="100%" border=0>
              <TR>
                            <TD vAlign=top background=../images/login_bg1.gif 
                bgColor=#f6f6f6 height=144> 
                              <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
                   <FORM action=../login/Login_Check.asp method=POST name=loginForm> 
                    <TR>
                      <TD width="36%" height=31>
                        <DIV align=center>��Ա����</DIV></TD>
                      <TD width="64%">
                                <INPUT id="username2" 
                        style="BORDER-RIGHT: #bbc0ca 1px solid; BORDER-TOP: #bbc0ca 1px solid; BORDER-LEFT: #bbc0ca 1px solid; BORDER-BOTTOM: #bbc0ca 1px solid; FONT-FAMILY: ����" 
                        size=13 name="UserID">
                              </TD></TR>
                    <TR>
                      <TD height=32>
                        <DIV align=center>�� �룺</DIV></TD>
                      <TD>
					            <INPUT id="password" 
                        style="BORDER-RIGHT: #bbc0ca 1px solid; BORDER-TOP: #bbc0ca 1px solid; BORDER-LEFT: #bbc0ca 1px solid; BORDER-BOTTOM: #bbc0ca 1px solid; FONT-FAMILY: ����" 
                        type=password size=13 name="Pwd">
                                 </TD></TR><INPUT 
                    type=hidden value=login name=send> <INPUT type=hidden 
                    value=false name=B_code> 
                    <TR>
                      <TD colSpan=2 height=29>
                        <DIV align=center><INPUT class=noo type=image 
                        src="../images/bbslogin.gif">
                                  &nbsp; <a href="../Reg/User_Reg.asp"><IMG 
                        src="../images/bbsreg.gif" border=0></a></DIV>
                              </TD></TR>
                                  <TR align="center"> 
                                    <TD height=21 colSpan=2>������Դ��Ϣ������������Ϣ</TD>
                                  </TR>
                    <TR>
                      <TD colSpan=2 height=26>
                        <DIV align=center>������Դ��Ϣ������������Ϣ</DIV></TD></TR></FORM></TABLE>
                            </TD>
                          </TR></TABLE></TD></TR></TABLE>
			</td>
        </tr>
      </table>
        
      <table width=95% 
                    border=0 cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>���¿���ר��</strong></font></td>
          <td width="130" height="22" valign="middle" align="center"> <a href="../Member/kyzx_add.asp"><font color="#FFFF00">��ѷ�������ר����Ϣ</font></a></td>
          <td width="38" height="22" valign="middle" align="right"><a href="../product/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></td>
        </tr>
        <tr> 
          <td colspan="3" valign="top" bgcolor="#FFFFFF" height="110"> 
            <%set rs_k=server.CreateObject("adodb.recordset")
		    rs_k.open "select top 10 * from zx_info where infotype='����ר��' and (city='����' or city2='����') order by infoid desc",conn,1,1
			
		  %>
            <table width="100%">
              <%if rs_k.bof and rs_k.eof then%>
              <tr> 
                <td height=23 align=left class=1>��ʱ����Ϣ</td>
              </tr>
              <%else%>
              <%k=0%>
              <tr> 
			  <%do while not rs_k.eof%>
                <td height=23 align=left class=1><a href="../pline/detail.asp?InfoID=<%=rs_k("infoid")%>" target="_blank"><%=rs_k("city")%>����<%=rs_k("city2")%></a></td>
              
              <%rs_k.movenext
			     k=k+1
				 if k mod 2=0 then
				 response.Write "</tr><tr>"
				 end if
			    loop
				rs_k.close
				set rs_k=nothing
			  %>
              <%end if%>
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
        
      <table width=95% 
                    border=0 cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>���»���ר��</strong></font></td>
          <td width="130" height="22" valign="middle" align="center"><a href="../Member/hyzx_add.asp"><font color="#FFFF00">��ѷ�������ר����Ϣ</font></a> 
          </td>
          <td width="38" height="22" valign="middle" align="right"><font color="#FFFFFF"><a href="../product/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></font></td>
        </tr>
        <tr> 
          <td colspan="3" valign="top" bgcolor="#FFFFFF" height="110"> 
            <%set rs_h=server.CreateObject("adodb.recordset")
					 rs_h.open "select top 10 * from zx_info where infotype='����ר��' and (city='����' or city2='����') order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='��Ӧ' Order by InfoID desc")
			  %>
            <table width="100%">
              <%if rs_h.bof and rs_h.eof then%>
              <tr> 
                <td height=23 align=center class=1>��ʱ����Ϣ</td>
              </tr>
              <%else%>
              <%j=0%>
              <tr> 
			  <%do while not rs_h.eof%>
                <td height=23 align=left class=1><a href="../fline/detail.asp?infoid=<%=rs_h("infoid")%>" target="_blank"><%=rs_h("city")%>����<%=rs_h("city2")%></a></td>
              <%rs_h.movenext
			   j=j+1
			   if j mod 2=0 then
			   response.Write "</tr><tr>"
			   end if
			    loop
				rs_h.close
				set rs_h=nothing
			  %>
              <%end if%>
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
        
      <table width=95% border=0 cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td height="20" valign="middle" bgcolor="#FF6600" width="150"><font 
            class=Ttitlew><strong><span class=Twhite14><img height=18 
            src="../images/yuanjian27.gif" 
            width=18></span></strong></font> <font color="#FFFFFF"><strong>���¼���������ҵ</strong></font></td>
          <td align="center" height="22" valign="middle"> <a href="../Reg/User_Reg.asp"><font color="#FFFF00">���ע���Ա</font></a></td>
          <td width="36" height="22" valign="middle" align="center"><font color="#FFFFFF"><a href="../company/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></font></td>
        </tr>
        <tr> 
          <td colspan="3" valign="top" bgcolor="#FFFFFF"> 
            <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 5 * from qyml where city='����' order by id desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='��Ӧ' Order by InfoID desc")
				%>
            <%if rs_info.bof and rs_info.eof then%>
            <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0>
              <tr> 
                <td  valign=top height=67>����Ϣ</td>
              </tr>
            </table>
            <%else %>
            <table cellspacing=2 cellpadding=2 border=0 width=100%>
              <%do while not rs_info.eof%>
              <tr align="center"> 
                <td><%=rs_info("qylb")%></td>
                <td><%=rs_info("city")%></td>
                <td><%=rs_info("qymc")%></td>
                <td><a href="../yellowpage/detail.asp?InfoID=<%=rs_info("ID")%>" target="_blank">����</a></td>
              </tr>
              <%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>
            </table>
            <%end if%>
          </td>
        </tr>
        </tbody> 
      </table>
    </TD>
    <TD vAlign=top align="right"> 
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
          <tbody> 
          <tr> 
            <td class=white12 valign=center bgcolor=#ff6600 height=22> 
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="9%"> 
                    <div align=center><img height=20 
                        src="../images/huo.gif" width=21></div>
                  </td>
                  
                <td width="22%"><font color=#ffffff>���»�Դ��Ϣ</font></td>
                  
                <td width="69%"> <font color="#FFFF00">������Ϣ���а�</font><font color="#FFFFFF"> 
                  | </font><a href="../Member/car_manage.asp"><font color="#FFFF00"></font></a> 
                  <a href="../Member/goods_add.asp"><font color="#FFFF00"> ��ѷ�����Դ��Ϣ</font></a><font color="#FFFFFF"> 
                  | </font><a href="../Member/goods_manage.asp"><font color="#FFFF00">�����Դ��Ϣ</font></a></td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr align=middle> 
            <td 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                valign=top height=140> 
                <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 10 * from info2 where infotype='��Դ��Ϣ' and (city='����' or city2='����')order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='��Ӧ' Order by InfoID desc")
				%>	
				<%if rs_info.bof and rs_info.eof then%>			
              <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top height=67>����Ϣ</td>
                </tr> 
              </table>
                <%else %>
<MARQUEE width="460" height="95" onmouseout=this.start() onmouseover=this.stop() direction="up" behavior="scroll" scrolldelay="500">		
            <table cellspacing=2 cellpadding=2 border=0 width=100%>
              <%do while not rs_info.eof%>
              <tr> 
                  
                <td  valign=top align="left"><%=rs_info("province")%>&nbsp;<%=rs_info("city")%>&nbsp;<%=rs_info("area")%></td>
				  
                <td  valign=top align="center">��</td>
				  
                <td  valign=top align="center"><%=rs_info("province2")%>&nbsp;<%=rs_info("city2")%>&nbsp;<%=rs_info("area2")%></td>
				  
                <td  valign=top align="center"><%=rs_info("cartype")%></td>
				  
                <td  valign=top align="center"><%=rs_info("carload")%>��</td>
				  
                <td  valign=top align="center"><%=rs_info("addtime")%></td>
				  
                <td  align="center"><A href="../good/detail.asp?InfoID=<%=rs_info("InfoID")%>" target="_blank">����</A></td>
                </tr> 
				<%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>				
              </table>
				 </MARQUEE>
			   <%end if%>			  
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="61%" height=24>
                  </td>
                  <td width="39%" height=24> 
                    <div align=center><a 
                        href="../truck/index.asp">����&gt;&gt;</a></div>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          </tbody> 
        </table>
        <TABLE cellSpacing=0 width="100%" border=0>
          <TBODY> 
          <TR> 
            <TD vAlign=top width="48%" height=151> 
              <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TBODY> 
                <TR> 
                  <TD class=white12 vAlign=center bgColor=#ff6600 height=22> 
                    <TABLE cellSpacing=0 width="100%" border=0>
                      <TBODY> 
                      <TR> 
                        <TD width="9%"> 
                          <DIV align=center><IMG height=20 
                        src="../images/huo.gif" width=21></DIV>
                        </TD>
                        
                      <TD width="22%"><FONT color=#ffffff>���¿ճ���Ϣ</FONT></TD>
                        
                      <TD width="69%"> <font color="#FFFF00">������Ϣ���а�</font><font color="#FFFFFF"> 
                        | </font><a href="../Member/car_manage.asp"><font color="#FFFF00"></font></a> 
                        <a href="../Member/car_add.asp"><font color="#FFFF00">��ѷ����ճ���Ϣ</font></a><font color="#FFFFFF"> 
                        | </font><a href="../Member/car_manage.asp"><font color="#FFFF00">����Դ��Ϣ</font></a></TD>
                      </TR>
                      </TBODY> 
                    </TABLE>
                  </TD>
                </TR>
                <TR align=middle> 
                  <TD 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                vAlign=top height=140> 
                    <TABLE height=67 cellSpacing=0 cellPadding=0 width="100%" 
                  border=0>
                      <TBODY> 
                      <TR> 
                        <TD  vAlign=top height=67>
                <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 10 * from info2 where infotype='��Դ��Ϣ' and (city='����' or city2='����') order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='��Ӧ' Order by InfoID desc")
				%>	
				<%if rs_info.bof and rs_info.eof then%>			
              <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top height=67>����Ϣ</td>
                </tr> 
              </table>
                <%else %>
<MARQUEE width="460" height="95" onmouseout=this.start() onmouseover=this.stop() direction="up" behavior="scroll" scrolldelay="500" class="p1">				
                        <table cellspacing=2 cellpadding=2 border=0 width=100%>
                          <%do while not rs_info.eof%>
                          <tr> 
                            <td  valign=top align="left"><%=rs_info("province")%><%=rs_info("city")%>&nbsp;<%=rs_info("area")%></td>
				            <td align="center">��</td>
				            <td align="center"><%=rs_info("province2")%><%=rs_info("city2")%>&nbsp;<%=rs_info("area2")%></td>
				            <td align="center"><%=rs_info("cartype")%></td>
				            <td align="center"><%=rs_info("carload")%>��</td>
				            <td align="center"><%=rs_info("addtime")%></td>
				            <td align="center"><A href="../truck/detail.asp?InfoID=<%=rs_info("InfoID")%>" target="_blank">����</A></td>
                </tr> 
				<%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>				
              </table>
				 </MARQUEE>
			   <%end if%>						
					    </TD>
                      </TR>
                      </TBODY> 
                    </TABLE>
                    <TABLE cellSpacing=0 width="100%" border=0>
                      <TBODY> 
                      <TR> 
                        <TD width="61%" height=24> 
                          <DIV align=center></DIV>
                        </TD>
                        <TD width="39%" height=24> 
                          <DIV align=center><A 
                        href="../good/index.asp">����&gt;&gt;</A></DIV>
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
      <DIV align=center> 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
          <tbody> 
          <tr> 
            <td class=white12 valign=center bgcolor=#ff6600 height=22> 
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="9%"> 
                    <div align=center><img height=20 
                        src="../images/huo.gif" width=21></div>
                  </td>
                  <td width="40%"><font color=#ffffff>���³�������</font></td>
                  <td width="51%"> 
                    <div align=right><font color="#FFFF00">��ѵǼǳ���</font></div>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr align=middle> 
            <td 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                valign=top height=140> 
              <table height=134 cellspacing=0 cellpadding=0 width="100%" 
                  border=0>
                <tbody> 
                <tr>
                  <td  valign=top>
                    <table cellspacing=1 cellpadding=2 border=0 width=100% bgcolor="#CCCC66">
                      <tr bgcolor="#99CC33"> 
                        <td>���ƺ�</td>
                        <td align="center">����</td>
                        <td  valign=top align="center">����</td>
                        <td  valign=top align="center">����</td>
                      </tr>
					  </table>				  </td>
                </tr>
                <tr> 
                  <td  valign=top height=100>
                <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 10 * from file_info a,qyml b where b.city='����' and a.gsid=b.user order by infoid desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='��Ӧ' Order by InfoID desc")
				%>	
				<%if rs_info.bof and rs_info.eof then%>			
              <table cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top>����Ϣ</td>
                </tr> 
              </table>
                <%else %>
<MARQUEE width="460" height="120" onmouseout=this.start() onmouseover=this.stop() direction="up" behavior="scroll" scrolldelay="500" class="p1">				
                    <table cellspacing=1 cellpadding=2 border=0 width=100% bgcolor="#CCCC66">
                      <%do while not rs_info.eof%>
                      <tr bgcolor="#FFFFFF"> 
                        <td><%=rs_info("carnum")%></td>
                        <td align="center"><%=rs_info("carpp")%></td>
                        <td  valign=top align="center"><%=rs_info("cartype")%></td>
                        <td  valign=top align="center"><A href="../cheliang/detail.asp?InfoID=<%=rs_info("InfoID")%>" target="_blank">����</A></td>
                      </tr>
                      <%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>
                    </table>
				 </MARQUEE>
			   <%end if%>				  </td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="61%" height=24> 
                    <div align=center></div>
                  </td>
                  <td width="39%" height=24> 
                    <div align=center><a 
                        href="../fline/index.asp">����&gt;&gt;</a></div>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          </tbody> 
        </table>
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
          <tbody> 
          <tr> 
            <td class=white12 valign=center bgcolor=#ff6600 height=22> 
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="9%"> 
                    <div align=center><img height=20 
                        src="../images/huo.gif" width=21></div>
                  </td>
                  <td width="40%"><font color=#ffffff>���»�Ա��Ϣ</font></td>
                  <td width="51%"> 
                    <div align=right></div>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr align=middle> 
            <td 
                style="BORDER-RIGHT: #ff9966 1px solid; BORDER-LEFT: #ff9966 1px solid; BORDER-BOTTOM: #ff9966 1px solid" 
                valign=top height=140> 
              <table height=67 cellspacing=0 cellpadding=0 width="100%" 
                  border=0>
                <tbody> 
                <tr> 
                  <td  valign=top height=67>
                <%  set rs_info=server.CreateObject("adodb.recordset")
				     rs_info.open "select top 10 * from qyml where city='����' order by id desc",conn,1,1
					'rs_info=conn.execute("select * from Info Where InfoType='��Ӧ' Order by InfoID desc")
				%>	
				<%if rs_info.bof and rs_info.eof then%>			
              <table height=67 cellspacing=0 cellpadding=0 width="100%" border=0> 
                <tr> 
                  <td  valign=top height=67>����Ϣ</td>
                </tr> 
              </table>
                <%else %>
<MARQUEE width="460" height="95" onmouseout=this.start() onmouseover=this.stop() direction="up" behavior="scroll" scrolldelay="500" class="p1">				
                    <table cellspacing=2 cellpadding=2 border=0 width=100%>
                      <%do while not rs_info.eof%>
                      <tr align="center"> 
                        <td><%=rs_info("qylb")%></td>
				        <td><%=rs_info("city")%></td>
				        <td><%=rs_info("qymc")%></td>
				        <td><%=rs_info("name")%></td>
				        <td><%=rs_info("phone2")%>-<%=rs_info("phone3")%></td>
				        <td><A href="../yellowpage/detail.asp?InfoID=<%=rs_info("ID")%>" target="_blank">����</A></td>
                </tr> 
				<%rs_info.movenext
			    loop
				rs_info.close
				set rs_info=nothing
				%>				
              </table>
				 </MARQUEE>
			   <%end if%>					  
				  </td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 width="100%" border=0>
                <tbody> 
                <tr> 
                  <td width="61%" height=24> 
                    <div align=center></div>
                  </td>
                  <td width="39%" height=24> 
                    <div align=center><a 
                        href="../yellowpage/index.asp">����&gt;&gt;</a></div>
                  </td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          </tbody> 
        </table>
      </DIV>
    </TD>
  </TR></TBODY></TABLE>
<table cellspacing=0 width=778 align=center border=0>
  <tbody> 
  <tr> 
      <td><img height=102 
      src="../images/banner2.gif" width=778 border=0></td>
  </tr>
  </tbody>
</table>
<table width="778" border="0" align="center" cellspacing="0">
  <tr> 
    <td width="545" valign="top">
      <table width=100% 
                    border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td width="118" height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>���¹�����Ϣ</strong></font></td>
          <td width="432" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF"><a href="../trade/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></font></div>          
          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF"  height="200">
		  <%set rs_gq=server.CreateObject("adodb.recordset")
		  rs_gq.open "select top 8 * from gq_info where city='����' order by infoid asc",conn,1,1
		  %>
		    <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <%if rs_gq.bof and rs_gq.eof then%>
              <tr>
              <td valign="top" colspan="4">��ʱ����Ϣ</td>
            </tr>
            <%else%>
            <%do while not rs_gq.eof%>
            <tr>
              <td width="72" valign="top" align="center"><%=rs_gq("infotype")%></td>
              <td width="152" valign="top" align="center"><a href="../trade/detail.asp?infoid=<%=rs_gq("infoid")%>" target="_blank"><%=rs_gq("title")%></a></td>
              <td width="215" valign="top"><%=rs_gq("province")%>&nbsp;<%=rs_gq("city")%>&nbsp;<%=rs_gq("area")%></td>
              <td width="98" valign="top"><%=rs_gq("addtime")%></td>
            </tr>
            <%rs_gq.movenext
			    loop
				rs_gq.close
				set rs_gq=nothing
			  %>
            <%end if%>
          </table></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td width="217" valign="top">
      <table width=100% 
                    border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>��Ƹ��Ϣ</strong></font></td>
          <td width="204" height="22" valign="middle">
            <div align="right"><font color="#FFFFFF"><a href="../company/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></font></div>
          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF" height="200">
		  <%set rs_z=server.CreateObject("adodb.recordset")
		   rs_z.open "select top 8 * from zp_info where infotype='��Ƹ��Ϣ' and city='����' order by infoid desc",conn,1,1
		  %>
            <table border="0" cellspacing="0" cellpadding="2" width="100%">
              <%if rs_z.bof and rs_z.eof then%>
              <tr> 
                <td colspan="3" valign="top">��ʱ����Ϣ</td>
              </tr>
			  <%else%>
			  <%do while not rs_z.eof%>
              <tr> 
                <td valign="top" align="center"> <a href="../job/detail.asp?infoid=<%=rs_z("infoid")%>" target="_blank"><%=rs_z("job")%></a></td>
                <td valign="top" align="center"><%=rs_z("province")%>&nbsp;<%=rs_z("city")%></td>
                <td valign="top" align="center"><%=rs_z("addtime")%></td>
              </tr>	
			  <%rs_z.movenext
			    loop
				rs_z.close
				set rs_z=nothing
			  %>
			  <%end if%>		  
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
</table>
<table width="778" border="0" align="center" cellspacing="0">
  <tr> 
    <td width="545" valign="top"> 
      <table width=100% 
                    border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>�����������</strong></font></td>
          <td width="204" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF"><a href="../Repairs/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></font></div>
          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF" height="200">
		  <%set file_info=server.CreateObject("adodb.recordset")
		    file_info.open "select top 8 * from repair_info where city='����' order by infoid desc",conn,1,1
		  %>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <%if file_info.bof and file_info.eof then%>
              <tr> 
                <td colspan="3" valign="top">��ʱ����Ϣ </td>
              </tr>
			  <%else%>
              <%do while not file_info.eof%>
			  <tr align="center"> 
                <td width="107" valign="top" align="center">
				<a href="../repairs/detail.asp?infoid=<%=file_info("infoid")%>" target="_blank"><%=file_info("title")%></a> </td>
                <td width="172" valign="top"><%=file_info("province")%>&nbsp;<%=file_info("city")%>&nbsp;<%=file_info("area")%></td>
                <td width="256" valign="top"><%=file_info("addtime")%></td>
                </tr>	
			  <%file_info.movenext
			    loop
				file_info.close
				set file_info=nothing
			  %>
			  <%end if%>		  
            </table>		  
		  
		  </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td width="217" valign="top"> 
      <table width=100% 
                    border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#FF6600">
        <tbody> 
        <tr bgcolor="#FF6600"> 
          <td width="419" height="20" valign="middle" bgcolor="#FF6600"><font color="#FFFFFF"><strong>�����˲�</strong></font></td>
          <td width="204" height="22" valign="middle"> 
            <div align="right"><font color="#FFFFFF"><a href="../company/index.asp" target="_blank"><font color="#FFFFFF">����</font></a></font></div>
          </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" bgcolor="#FFFFFF" height="200"> 
		  <%set rs_q=server.CreateObject("adodb.recordset")
		    rs_q.open "select top 8 * from zp_info where infotype='��ְ��Ϣ' and city='����' order by infoid desc",conn,1,1
		  %>
            <table border="0" cellspacing="0" cellpadding="2" width="100%">
              <%if rs_q.bof and rs_q.eof then%>
              <tr> 
                <td colspan="3" valign="top">��ʱ����Ϣ </td>
              </tr>
			  <%else%>
              <%do while not rs_q.eof%>
			  <tr> 
                <td valign="top" align="center"> <a href="../job/detail.asp?infoid=<%=rs_q("infoid")%>" target="_blank"><%=rs_q("job")%></a> 
                </td>
                <td valign="top" align="center"><%=rs_q("province")%>&nbsp;<%=rs_q("city")%></td>
			    <td valign="top" align="center"><%=rs_q("addtime")%></td>
			  </tr>	
			  <%rs_q.movenext
			    loop
				rs_q.close
				set rs_q=nothing
			  %>
			  <%end if%>		  
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
</table>
<TABLE cellSpacing=0 width=778 align=center border=0>
  <TBODY>
  <TR>
    <TD><img height=113 
      src="../images/bannar.jpg" width=778 border=0></TD>
  </TR></TBODY></TABLE>
<table width="778" border="0" align="center" cellspacing="0">
  <tr> 
    <td width="554" valign="top"> 
      <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td width="48%" height="151" valign="top"> 
            <table cellspacing=0 cellpadding=0 width=100% border=0>
              <tbody> 
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        <div align="center"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">������̬</font>��������������<a href="../user/login.asp" target="_blank"></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                <td height="127" valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
                  <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
          <tbody> 
          <%set rs_n=conn.execute("select top 8 * from NewsData where NClassID=1 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
            <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="../News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a> 
              (<%=month(rs_n("RegTime"))%>/<%=day(rs_n("RegTime"))%>)</td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
          </tbody> 
        </table>
        
      </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td width="1%" valign="top">&nbsp;</td>
          <td width="51%" valign="top"> 
            <table cellspacing=0 cellpadding=0 width=100% border=0>
              <tbody> 
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        <div align="center"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">����֪ʶ</font>��������������<a href="../user/login.asp" target="_blank"></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                <td height="127" valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
                  <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
          <tbody> 
          <%set rs_n=conn.execute("select top 8 * from NewsData where NClassID=2 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
            <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="../News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a> 
              (<%=month(rs_n("RegTime"))%>/<%=day(rs_n("RegTime"))%>)</td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
          </tbody> 
        </table>
        
      </td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="220" valign="top"> 
      <table width=217 height="202" border=0 align="right" cellpadding=0 cellspacing=0>
        <tbody> 
        <tr> 
          <td height="22" bgcolor="#FF6600"><font face=���� size=3> ��<font color="#FFFFFF" size="2"><strong>��������</strong></font></font> 
          </td>
        </tr>
        <tr> 
          <td height="184" valign="top" 
          style="BORDER-RIGHT: #cccccc 1px solid; PADDING-RIGHT: 8px; BORDER-TOP: #cccccc 3px solid; PADDING-LEFT: 8px; FONT-SIZE: 12px; PADDING-BOTTOM: 10px; BORDER-LEFT: #cccccc 1px solid; LINE-HEIGHT: 150%; PADDING-TOP: 10px; BORDER-BOTTOM: #cccccc 1px solid"> 
            <table width="100%" height="80" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hntrans.com:81/gs/xinlichsele.php" target="_blank" class="blacklink">��̲�ѯ</a>]</td>
                <td height="22" align="center">��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.mapbar.com:8000/commonmapright/?title=�㽭������&cityCode=0571&width=760&zm=20&logo=true&cityOption=ture&mapstyle=0&state=map666&advscrolling=false&color=imgblue&headType=lydt&headPic=false" target="_blank" class="blacklink">������ͼ</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://cb.kingsoft.com/" target="_blank" class="blacklink">Ӣ������</a>]</td>
                <td height="22" align="center">��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://zj.183.com.cn/quire/code.htm" target="_blank" class="blacklink">�ʱ��ѯ</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hlbr-l-tax.gov.cn/xxfw/slcx/001%5B1%5D.htm" target="_blank" class="blacklink">˰�ʲ�ѯ]</a></td>
                <td height="22" align="center">��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://info./jinchu/whhx/default.shtml" target="_blank">������</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.hzti.com/" target="_blank" class="blacklink">��ͨ״��</a>]</td>
                <td height="22" align="center"> ��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.xichang.tv/calendar.htm" target="_blank" class="blacklink">��������</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.5l.com.cn/second/shouce/kgdm.htm" target="_blank">���ʿո�</a>]</td>
                <td height="22" align="center">��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.5l.com.cn/second/shouce/gonglushoufei.htm" target="_blank" class="blacklink">��·�շ�</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="../tool/sjhb.htm" target="_blank">�������</a>]</td>
                <td height="22" align="center"> ��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://info./entr/list_cgs.shtml" target="_blank">����׷��</a>]</td>
              </tr>
              <tr>
                <td height="22" align="center"><a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="../tg/bgdm.asp" target="_blank">���ش���</a>]</td>
                <td height="22" align="center">��<a href="http://www.t7online.com/China.htm" 
      target=_blank>[</a><a href="http://www.chemaid.com/sample/user.asp" target="_blank" class="blacklink">Σ��Ʒ��</a>]</td>
              </tr>
            </table>
                </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
</table>
<table width="778" border="0" align="center" cellspacing="0">
  <tr> 
    <td width="554" valign="top"> 
      <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td width="48%" height="151" valign="top"> 
            <table cellspacing=0 cellpadding=0 width=100% border=0>
              <tbody> 
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        <div align="center"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">��ҵ����</font>��������������<a href="../user/login.asp" target="_blank"></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                <td height="127" valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
                  <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
          <tbody> 
          <%set rs_n=conn.execute("select top 8 * from NewsData where NClassID=3 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
            <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="../News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a> 
              (<%=month(rs_n("RegTime"))%>/<%=day(rs_n("RegTime"))%>)</td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
          </tbody> 
        </table>
      </td>
              </tr>
              </tbody> 
            </table>
          </td>
          <td width="1%" valign="top">&nbsp;</td>
          <td width="51%" valign="top"> 
            <table cellspacing=0 cellpadding=0 width=100% border=0>
              <tbody> 
              <tr> 
                <td height=22 valign="middle" bgcolor="#FF6600" class=white12> 
                  <table width="100%" border="0" cellspacing="0">
                    <tr> 
                      <td width="11%"> 
                        <div align="center"></div>
                      </td>
                      <td width="89%"><font color="#FFFFFF">������չ</font>��������������<a href="../user/login.asp" target="_blank"></a></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align=middle> 
                <td height="127" valign="top" 
          style="BORDER-RIGHT: #FF9966 1px solid; BORDER-LEFT: #FF9966 1px solid; BORDER-BOTTOM: #FF9966 1px solid"> 
                  <table 
      style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; BORDER-LEFT: #cccccc 1px solid; BORDER-BOTTOM: #cccccc 1px solid" 
      cellspacing=0 cellpadding=1 width=100% border=0>
          <tbody> 
          <%set rs_n=conn.execute("select top 8 * from NewsData where NClassID=4 Order By NewsID Desc")
		  if rs_n.eof and rs_n.bof then
		%>
          <tr> 
            <td height="20">&nbsp;<font face='Wingdings'>w</font>&nbsp;No Data!</td>
          </tr>
          <%  
		  else
		  do while not rs_n.eof
		  topic=gotTopic(rs_n("Title"),MaxLen)
          topic=replace(server.HTMLencode(topic)," ","&nbsp;")
          topic=replace(topic,"'","&nbsp;")
		  %>
          <tr> 
            <td height="22">&nbsp;<font face='Wingdings'>w</font>&nbsp;<a href="../News/NewsDetail.asp?NewsID=<%=rs_n("NewsID")%>" target="_blank"><%=topic%></a> 
              (<%=month(rs_n("RegTime"))%>/<%=day(rs_n("RegTime"))%>)</td>
          </tr>
          <%
		  rs_n.movenext
		  loop
		  end if
		  rs_n.close
		  set rs_n=nothing
		  %>
          </tbody> 
        </table>
      </td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="220" valign="top"> 
      <table width=217 height="202" border=0 align="right" cellpadding=0 cellspacing=0>
        <tbody> 
        <tr> 
          <td height="22" bgcolor="#FF6600"><font face=���� size=3> ��<font color="#FFFFFF" size="2"><strong>���񹤾�</strong></font></font> 
          </td>
        </tr>
        <tr> 
          <td height="180" valign="top" 
          style="BORDER-RIGHT: #cccccc 1px solid; PADDING-RIGHT: 8px; BORDER-TOP: #cccccc 3px solid; PADDING-LEFT: 8px; FONT-SIZE: 12px; PADDING-BOTTOM: 10px; BORDER-LEFT: #cccccc 1px solid; LINE-HEIGHT: 150%; PADDING-TOP: 10px; BORDER-BOTTOM: #cccccc 1px solid"> 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
              <tbody> 
              <tr align="center"> 
                <td height=20><a href="http://www.t7online.com/China.htm" 
      target=_blank>[����Ԥ��]</a></td>
                <td height="20"><a href="http://www.21page.net/public/code_country.asp" target="_blank">[�ʱ�����</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a 
                        href="http://www.go2map.com" target="_blank">�й���ͼ</a>]</td>
                <td height="20"><a href="../tool/rili.htm" target="_blank">[�� �� ��]</a></td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a 
                        href="http://www.id5.cn/uindex.jsp?webid=cx56w" target="_blank">��ݺ˲�</a>]</td>
                <td height="20">[<a href="http://www.hntrans.com:81/gs/xinlichsele.php" target="_blank">��̲�ѯ</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a href="http://train.chinamor.cn.net/cctt.asp" target="_blank">�г�ʱ��</a>]</td>
                <td height="20">[<a 
                        href="http://www.123cha.com" target="_blank">�ֻ���ѯ</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a href="http://www.21page.net/public/time_world.asp" target="_blank">����ʱ��</a>]</td>
                <td height="20">[<a href="../tool/road.htm" 
      target=_blank>������ѯ</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a href="../tool/meipay.asp" target="_blank"><font color="#FF0000">��������</font></a>]</td>
                <td height="20">[<a class=blue 
                  href="http://ent.tom.com/tv/forecast/2.html" 
                  target=_blank>���ӽ�Ŀ</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a class=blue 
                  href="http://act.it.sohu.com/download/download.php" 
                  target=_blank>װ�����</a>]</td>
                <td height="20">[<a class=blue 
                  href="http://www.chinaren.com/s2005/topmusic.shtml" 
                  target=_blank>��������</a>]</td>
              </tr>
              <tr align="center"> 
                <td height=20>[<a 
                        href="http://www.alexa.com/site/ds/top_sites?ts_mode=lang&lang=zh_gb2312" target="_blank" class=titlefont>������ѯ</a>]</td>
                <td height="20">[<font color="#FFFFFF"><font color="#FFFFFF"><font color="#FFFFFF"><a href="http://site.cx56w.com" target="_blank"><font color="#FF0000">�������</font></a></font></font></font>]</td>
              </tr>
              </tbody> 
            </table>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
</table>
<table border="0" width="778" id="table2" height="80" cellpadding="2" align="center" cellspacing="1" bgcolor="#ff6600">
  <tr> 
    <th bgcolor="#FF9900" width="22" align="center"> 
      <p align="right"><font size="3" color="#FFFFFF">���ط�վ</font></p>
    </th>
    <td bgcolor="#FFFFFF"> 
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FF9900">
        <!--DWLayoutTable-->
        <tr align="center" valign="middle"> 
    <td width="69" height="23" bgcolor="#FFFFFF"><a href="beijing.asp" target="_blank">��������</a></td>
    <td width="69" bgcolor="#FFFFFF"><a href="shanghai.asp" target="_blank">�Ϻ�����</a></td>
    <td width="69" bgcolor="#FFFFFF"><a href="tianjin.asp" target="_blank">������</a></td>
    <td width="69" bgcolor="#FFFFFF"><a href="chongqing.asp" target="_blank">�������</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="shangdong.asp" target="_blank">ɽ������</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="shanxi.asp" target="_blank">ɽ������</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="henan.asp" target="_blank">���Ϸ���</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="hebei.asp" target="_blank">�ӱ�����</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="hunan.asp" target="_blank">���Ϸ���</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="hubei.asp" target="_blank">��������</a></td>
    <td width="70" bgcolor="#FFFFFF"><a href="hainan.asp" target="_blank">���Ϸ���</a></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td height="23" align="center" valign="middle" bgcolor="#FFFFFF"><a href="jiangxi.asp" target="_blank">��������</a></td>
    <td bgcolor="#FFFFFF"><a href="sichuan.asp" target="_blank">�Ĵ�����</a></td>
    <td bgcolor="#FFFFFF"><a href="anhui.asp" target="_blank">���շ���</a></td>
    <td bgcolor="#FFFFFF"><a href="fujian.asp" target="_blank">��������</a></td>
    <td bgcolor="#FFFFFF"><a href="jiangsu.asp" target="_blank">���շ���</a></td>
    <td bgcolor="#FFFFFF"><a href="guangxi.asp">��������</a></td>
    <td bgcolor="#FFFFFF"><a href="guangdong.asp">�㶫����</a></td>
    <td bgcolor="#FFFFFF"><a href="guizhou.asp" target="_blank">���ݷ���</a></td>
    <td bgcolor="#FFFFFF"><a href="yunnan.asp" target="_blank">���Ϸ���</a></td>
    <td bgcolor="#FFFFFF"><a href="zhejiang.asp" target="_blank">�㽭����</a></td>
    <td bgcolor="#FFFFFF"><a href="shan xi.asp" target="_blank">��������</a></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td height="25" bgcolor="#FFFFFF"><a href="gansu.asp" target="_blank">�������</a></td>
    <td bgcolor="#FFFFFF"><a href="qinghai.asp" target="_blank">�ຣ����</a></td>
    <td bgcolor="#FFFFFF"><a href="liaoning.asp" target="_blank">��������</a></td>
    <td bgcolor="#FFFFFF"><a href="heilongjiang.asp" target="_blank">��������</a></td>
    <td bgcolor="#FFFFFF"><a href="jilin.asp" target="_blank">���ַ���</a></td>
    <td bgcolor="#FFFFFF"><a href="ningxia.asp" target="_blank">���ķ���</a></td>
    <td bgcolor="#FFFFFF"><a href="neimenggu.asp" target="_blank">���ɷ���</a></td>
    <td bgcolor="#FFFFFF"><a href="xinjiang.asp" target="_blank">�½�����</a></td>
    <td bgcolor="#FFFFFF"><a href="xizang.asp" target="_blank">���ط���</a></td>
    <td bgcolor="#FFFFFF"><a href="../index.asp" target="_blank">��������</a></td>
    <td bgcolor="#FFFFFF"><a href="http://www.cx56w.com">������վ</a></td>
  </tr>
</table>
    </td>
  </tr>
</table>
<TABLE width="778" border="0" cellspacing="0" cellpadding="2" align="center">
  <TBODY> 
  <TR> 
    <TD align=center height="10"></TD>
  </TR>
  <TR> 
    <TD align=center height="3" bgcolor="#FF6600"></TD>
  </TR>
  <TR> 
    <TD height="26" align=center bgcolor="#F7F7F7"><a 
      href="../aboutus/index.asp">��������</a> | <a 
      href="../aboutus/service.asp" target=_blank>��������</a> | <a 
      href="../aboutus/guide.asp">��վָ��</a> | <a 
      href="../aboutus/hyzh.asp">��Ա�ʷ�</a> | <a 
      href="../aboutus/exp.asp">֧����ʽ</a> |��<a 
      href="../aboutus/ad.asp">������</a> | <a 
      href="../aboutus/cooperate.asp">�г�����</a> | <a 
      href="../aboutus/supply.asp">������ר��</a></TD>
  </TR>
  </TBODY> 
</TABLE>
<table cellspacing=2 cellpadding=2 width=778 align=center border=0>
  <tbody> 
  <tr> 
    <td valign=center align=center>CopyRight &copy; 2003-2006 <a href="http://www.cx56w.com" target="_blank">Www.Cx56w.Com</a> 
      All Rights Reserved.<nobr></nobr> </td>
  </tr>
  <tr> 
    <td valign=center align=center>�������ߣ�13116601816 �������䣺info@cx56w.com</td>
  </tr>
  <tr> 
    <td valign=center align=center>����ʵ�������������� �����ţ���ICP��05059371��</td>
  </tr>
  </tbody> 
</table>
</BODY></HTML>
