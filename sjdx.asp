<!--#include file="Inc/Conn2.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
const MaxNumber=50
const MaxNum=8
const MaxLen=30
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
<%typeselect=request("type")
%>
<%
function title_left(str,num)
if len(str)>num then
response.write left(str,num)&""
else
response.write str
end if
end function
%>
<HTML>
<HEAD>
<TITLE>168������Ϣ��--ȫ����רҵ����ŵ�������Ϣ����ƽ̨|����,������,�й�������,������Ϣ��</TITLE>
<meta name=keywords content="����,������,�й�������,168������Ϣ��">
<meta name=description content="����-168������Ϣ���Ǵ��͹���������Ѷ��������ƷӦ��,��Ʒ���ҽ��ܵ�����ý��.��ӭ��Ϊ168������Ϣ��ע���Ա�ṩ�������!Find the logistics products and services in this website!">
<LINK href="images/css.css" type=text/css rel=stylesheet>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content=. name=description>
<META  content=��>
<SCRIPT language=JavaScript type=text/JavaScript>
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
function openwindow(url,width,height) { 	left1 =0; 	top1 = 0; 	window.open(url,"","width=" + width + ",height=" + height + ",toolbar=no,menubar=no,scrollbars=yes,location=no,status=yes,left=" + left1.toString() + ",top=" + top1.toString()); }
//resizable=yes, ���
//-->
</SCRIPT>
<script language=JavaScript src="inc/p_c_a.js"></script>
<META content="MSHTML 6.00.2462.0" name=GENERATOR>
<style type="text/css">
<!--
.STYLE1 {color: #FFFFFF}
-->
</style>
</HEAD>
<BODY topMargin=0>
<!--#include virtual="sjtop.asp" -->
<table cellspacing=0 cellpadding=0 width=768 align=center border=0>
  <tbody> 
  <tr> 
    <td background=images/backdropi.gif bgcolor=#ffffff> 
      <table cellspacing=0 cellpadding=0 width="96%" align=center border=0>
        <tbody> 
        <tr> 
          <td><img height=100 src="images/sjbt.gif" 
        width=760></td>
        </tr>
        <tr> 
          <td background=images/fz02.gif> 
            <div align=center> 
              <table style="TEXT-INDENT: 25px; LINE-HEIGHT: 150%" cellspacing=1 
            cellpadding=7 width="96%" bgcolor=#d9d9d9 border=0>
                <tbody> 
                <tr bgcolor=#eeffdb> 
                  <td valign=top width="30%" 
                  height=88>168������Ϣ����ͨ��֧���ֻ�,С��ͨ����<br>
                    ˫�򻥶�,�����ȫ��֧�ֵĶ����ط��ţ� 
                    <p>�������168������Ϣ���Ļ�Ա��������ע��ʱ��д����ȷ���ֻ��Ż�С��ͨ�ţ����������Ļ�Ա�ʻ���Ԥ���˶����ʷѣ��Ϳ���ͨ�����Ų�ѯ�����Ϣ��<br>
                      <img 
                  height=77 src="images/fsgz.gif" width=214></p>
                    <p>������168������Ϣ����վ���վ���л�����ͨ�з����˻�Դ��ճ���Ϣ������Ч�������������Ա������һ����������Ϣ��ƥ�����Ϣ��������ֻ����յ�������Ϣ�� 
                    <p>���磬����������վ�Ϸ�����һ���ӳ����������Ļ�Դ��Ϣ����Ч��Ϊ3�죬��ô�ڽ���3���ڣ������������Ա�����˴ӳ����������ĳ�Դ��Ϣ����168������Ϣ�����Զ���������Ϣ��������ֻ��ϣ������ο��� 
                      <br>
                      <font color="#000080">������������г�25�ֺ�����վ �ź�ΰ 022-4724556 
                      13545238688 [�������288��]��168������Ϣ��</font> </p>
                  </td>
                  <td valign=top> 
                    <p><strong>��ѯ��Դ�༭����</strong><br>
                      �����������ţ���010���Ǻţ���������������ţ���021��C <br>
                      ���磬���л�Ҫ�ӱ��������Ϻ������ѯ�������Ϻ��ĳ�Դ���������¶��ţ�<font color=#ff0000><b> 
                      <font size=2>010*021C</font></b></font> 
                    <p><strong>��ѯ��Դ�༭����</strong><br>
                      �����������ţ���0379���Ǻţ���������������ţ���0431��H&nbsp; <br>
                      ���磬���г������ѯ��֣�����������Ļ�Դ���������¶��ţ�<font color=#ff0000><b><span 
                  class=style10><font 
                  size=2>0371*0431H</font></span></b></font><font size=2> </font><br>
                      �ƶ��û����͵� <b><font size="2" color="#800080">620127870290</font></b><br>
                      ��ͨ�û����͵� <b><font size="2" color="#800080">620127870290</font></b><br>
                    </p>
                    <p>����������ѯ����,��������Ч���ڵ���Ϣ����168������Ϣ�����������������¶��ŵ�����ֻ�/С��ͨ�ϣ� <br>
                      <font 
                  color=#000080>�����������Ϻ������г�30�ֺ������ ������01062500559 13114861250[����2����Ϣ] 
                      [�������288��]��168������Ϣ��<br>
                      ���ݵ��人�ػ�25�־������� ֣��0571-85953056 13642182516����[����5����Ϣ][�������267��]��168������Ϣ��</font><br>
                      <font 
                  color=#000080>����[����5����Ϣ]</font> ��������5�������ѯ��������Ϣ���������ٴη���ͬ���Ĳ�ѯ���ţ�168������Ϣ����������һ����Ϣ���㣮 
                      <br>
                      ������������ֻ������Ĳ�ѯ���Ÿ�ʽ���ԣ���û�������ѯ��������Ϣ����168������Ϣ�������ͣ��Խ�ʡ����168������Ϣ��Ԥ��Ķ����ʷѣ�</p>
                    </td>
                </tr>
                <tr bgcolor=#eeffdb> 
                  <td width="30%" height=88><img height=77 
                  src="images/zfbz.gif" width=214></td>
                  <td valign=top>�������ֻ�/С��ͨ���Ͳ�ѯ����ʱ�����㷢����������һ�������Ų��Ž��۳�0.1Ԫ�� <br>
                    ������168������Ϣ���������ֻ�/С��ͨ���������Ϣʱ��168������Ϣ�������Ļ�Ա�ʻ���Ԥ��Ķ����ʷ��п۳�0.1Ԫ�� 
                    <br>
                    &nbsp;&nbsp;&nbsp; ��������ʱ�ڻ�Ա���Ļ��л�����ͨ��鿴168������Ϣ�����㷢��������ŵļ�¼��</td>
                </tr>
                </tbody> 
              </table>
            </div>
          </td>
        </tr>
        <tr> 
          <td><img height=22 src="images/fz03.gif" 
        width=760></td>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<tr> 
  <td align="center"> 
    <!--#include file="sjbottom.html" -->
    <!-- Live800���߿ͷ�ͼ��:cx56w1[������] ��ʼ-->
    <script language="javascript" src="http://chat.live800.com:8080/live800/chatClient/floatButton.js?jid=9756223015&companyID=32913&configID=24234&codeType=custom"></script>
    <!-- Live800���߿ͷ�ͼ��:cx56w1[������] ����-->
    <!-- Live800���߿ͷ����ܴ��� ��ʼ-->
    <script language="javascript" src="http://chat.live800.com:8080/live800/chatClient/monitor.js?jid=9756223015&companyID=32913&configID=23583&codeType=custom"></script>
    <!-- Live800���߿ͷ����ܴ��� ����-->
</BODY>
</HTML>
