<%@language=vbscript codepage=936 %>
<%
option explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<HTML>
<HEAD>
<TITLE>��Աע��ȷ��ҳ��-����������</TITLE>
<META content=False name=vs_snapToGrid>
<META content=True name=vs_showGrid>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META 
content=����������:ȫ��������������������վ���ṩ����Ʒ��������Ʒ��ҵ��������Ʒ������Ʒ��Ŀ������Ʒ���š�����Ʒ���顢����Ʒ��չ����Ϣ��רҵ����Ʒ��վ�� 
name=description>
<META 
content=ȫ������B2B����ƽ̨��������ǧ��������������������21���������ȷ桢����Ʒ��������Ʒ���մ�����������������̩����Ʒ�����������������ᡢ����ľ��ʯ������������Ʒ��������觡���ľ�ڼ�����ʯ�������������. 
name=keywords>
<META content=TRUE name=MSSmartTagsPreventParsing>
<META http-equiv=MSThemeCompatible content=Yes><LINK href="images/css.css" 
type=text/css rel=stylesheet>
<script language="JavaScript"> 
var isOK;
function fCheck(){
	/*something wrong*/
	/*form.UserName.trim();*/
	if( document.form.userid.value =="") {
		alert("\�����������û��� !");
		document.form.userid.focus();
		return false;
	}
	if( document.form.userid.value.length<5||document.form.userid.value.length>20) {
		alert("\�����û�������Ӧ����5��20���ַ�֮��!");
		document.form.userid.focus();
		return false;
	}
	if ( fIsNumber(document.form.userid.value.charAt(0),"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")!=1 ){
		alert("\�����û���ֻ������ĸ��ͷ!");
		document.form.userid.focus();
		return false;
	}
	if ( fIsNumber(document.form.userid.value,"1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ._-")!=1 ){
		alert("\�����û���Ӧ�������֡���ĸ,��������ֺ��ֵ������ַ�!");
		document.form.userid.focus();
		return false;		
	}
	 var mystring = "a" + document.form.userid.value;
       if (mystring.indexOf("falun")>0 ){
               alert("��ʹ���˷Ƿ����û���,������ѡ��һ��");
               return false;
       }
	return true;
}
//****�ж��Ƿ���Number.
function fIsNumber (sV,sR) {
	var sTmp;
	if(sV.length==0){ return (false);}
	for (var i=0; i < sV.length; i++){
		sTmp= sV.substring (i, i+1);
		if (sR.indexOf (sTmp, 0)==-1) {return (false);}
	}
	return (true);
}
//***
</script>
<STYLE type=text/css>.p14 {
	FONT-SIZE: 14px; COLOR: #000000; LINE-HEIGHT: 130%
}
.S {
	FONT-SIZE: 12px
}
.style16 {
	FONT-WEIGHT: bold
}
.button {
	BORDER-RIGHT: #ffffff thin; BACKGROUND-POSITION: center center; BORDER-TOP: #ffffff thin; FONT-SIZE: 12px; BORDER-LEFT: #ffffff thin; WIDTH: 50px; COLOR: #ffffff; BORDER-BOTTOM: #ffffff thin; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 20px; BACKGROUND-COLOR: #999999
}
.button01 {
	BORDER-RIGHT: #9c9a9c 1px solid; BACKGROUND-POSITION: center center; BORDER-TOP: #9c9a9c 1px solid; FONT-SIZE: 12px; MARGIN: 1px; BORDER-LEFT: #9c9a9c 1px solid; WIDTH: 100px; COLOR: #666666; BORDER-BOTTOM: #9c9a9c 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: auto; BACKGROUND-COLOR: #cccccc
}
.button02 {
	BORDER-RIGHT: #9c9a9c 1px solid; BACKGROUND-POSITION: center center; BORDER-TOP: #9c9a9c 1px solid; FONT-SIZE: 12px; MARGIN: 1px; BORDER-LEFT: #9c9a9c 1px solid; WIDTH: 50px; COLOR: #666666; BORDER-BOTTOM: #9c9a9c 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 20px; BACKGROUND-COLOR: #cccccc
}
.button04 {
	BORDER-RIGHT: #aaaaaa 1px solid; BORDER-TOP: #fefefe 1px solid; FONT-SIZE: 12px; BACKGROUND: #ececec; BORDER-LEFT: #fefefe 1px solid; WIDTH: 380px; COLOR: #000000; PADDING-TOP: 1px; BORDER-BOTTOM: #aaaaaa 1px solid; WHITE-SPACE: normal; HEIGHT: 23px
}
.button03 {
	BORDER-RIGHT: #aaaaaa 1px solid; BORDER-TOP: #fefefe 1px solid; FONT-SIZE: 12px; BACKGROUND: #ececec; BORDER-LEFT: #fefefe 1px solid; COLOR: #000000; PADDING-TOP: 1px; BORDER-BOTTOM: #aaaaaa 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 23px
}
.style17 {
	FONT-WEIGHT: bold; COLOR: #cc3333
}
</STYLE>
<META content="MSHTML 6.00.2462.0" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginwidth="0" marginheight="0">
<table cellspacing=0 cellpadding=0 width=776 align=center border=0>
  <tbody> 
  <tr> 
    <td valign=bottom> 
      <table cellspacing=0 cellpadding=0 width=776 border=0>
        <tbody> 
        <tr> 
          <td> 
            <table style="BORDER-BOTTOM: #5286b5 1px solid" height=70 cellspacing=0 
      cellpadding=0 width=776 border=0>
              <tbody> 
              <tr> 
                <td width=265 height=64><a href="../"><img height=65 
            alt=������ҳ src="images/logo.gif" width=265 align=bottom 
            border=0></a></td>
                <td valign=bottom align=middle width=511> 
                  <table height=50 cellspacing=0 cellpadding=0 width=489 border=0>
                    <tbody> 
                    <tr> 
                      <td width=180><img height=50 src="images/title1_re1.gif" 
                  width=180></td>
                      <td width=180><img height=50 src="images/title1_re2.gif" 
                  width=180></td>
                      <td width=129><img height=50 src="images/title_re3.gif" 
                  width=129></td>
                    </tr>
                    </tbody> 
                  </table>
                </td>
              </tr>
              </tbody> 
            </table>
            <table cellspacing=0 cellpadding=10 width="100%" border=0>
              <tbody> 
              <tr> 
                <td align=middle width="12%"><img height=30 
                  src="images/htxt_01.gif" 
        width=500></td>
                <td width="88%"><span class=S><span 
            class=style17>��<font 
      color=#ff0000><strong>24Сʱ�ͷ����ߣ�(0)13116601816</strong></font></span></span></td>
              </tr>
              </tbody> 
            </table>
            
          </td>
        </tr>
        </tbody> 
      </table>
      
    </td>
  </tr>
  </tbody> 
</table>
<table cellspacing=0 width=768 border=0 align="center">
  <tbody> 
  <tr> 
    <td colspan=2><img height=80 
      alt=ֻҪע���˱�վ�Ļ�Ա�����ͻ�õ��������Լ�����վ���������Լ�����վ�Ϸ����ճ���Ϣ����Դ��Ϣ�������������û��������������ҵ��㡣 
      src="images/56hy.gif" width=776></td>
  </tr>
  </tbody> 
</table>
<table id=Table2 cellspacing=0 cellpadding=0 width=759 border=0 align="center">
  <tbody> 
  <tr> 
    <td valign=top height=7><img height=7 src="images/m_15.gif" 
  width=778></td>
  </tr>
  <tr> 
    <td align=middle background=images/m_13.gif height=60><img height=41 
      alt=��Աע���һ������Ա�������� src="images/m_03.gif" width=639></td>
  </tr>
  <tr> 
    <td align=middle background=images/m_13.gif> 
      <table width="96%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td> 
            <p><br>
              <span class="p14hon">ʹ��168������Ϣ����ȫ�����ط������ṩ�ķ�����Ҫ�������·�������</span><br>
              <br>
              1�� 168������Ϣ��ƽ̨��Ʒ��Ʒ�ơ�VI���̱�Ⱦ�Ϊ�������������У�δ����ɣ��κε�λ����˲� �����á��Ƿ����á��Լ��κ���ʽ���޸�����֪ʶ��Ȩ�� 
              <br>
              <br>
              2�� ����������վ���û�Ӧ�����л����񹲺͹���صķ��ɺͷ��棬���س���������������򣬲������κη��������Լ����л�ɫ�������ݵĴ����� 
              ����һ�к���ɷ������Ը���168������Ϣ�����е��κη������Σ������������Ὣ�����˵�һ�������Ϣ�ṩ���������ء� <br>
              <br>
              3�������������û�Ӧ�ṩ׼ȷ��ע�����ϡ������û����ϲ�ȫ��׼ȷ����ɵķ����жϻ�ȱ�ݻ������κ���ʧ�ͺ�������������������κ����Ρ� 
              <br>
              <br>
              4�� �����������������Թ����û���˽���ϣ��������ע����Ϣ��������ҵ��Ϣ���������ơ���ϵ�绰��ͨѶ��ʽ�ȿ������û���ɣ�����������ʽע���û���������������������Ȩ�Ե����ʼ��򳣹��ʼ���ʽΪ���û��ṩ����������Ϣ�����������û���ɡ� 
              <br>
              <br>
              5�� �������������ṩ���������ʻ�������վ����Դ�������ӡ��������ӵ�������վ����Դ�Ƿ�������ã��������������赣������ʹ�û�����������վ����Դ��������ʧ���𺦣���վҲ�������κ����Ρ�<br>
              <br>
              6�� �����������������ݵİ�Ȩ��ԭ�����ߺ���������ͬ���У��κθ��ơ�ת�غ�������Ϊ��Ӧ����ԭ�����ߺ�������������Ȩ�� <br>
              <br>
              7�� �����������û���������Ϣ��������������У��κθ��ơ�ת����Ϊ������õ�������������������Ȩ��δ������Ȩ���κ��û��������κ���ʽ���κ��ֶν��κ���Ϣת�����������磬��������Ȩ��ʱ��ֹ�û��ķ���Э�飬�������շ��øŲ������� 
              ���������ɴ����Ȩ���� <br>
              <br>
              8�� δ��������ɣ��κ��û��������ñ���������������Ϣ����κι����Ϣ��������Ȩ��ʱ��ͣ�û�������������������Ȩ��ֹ�û������������շ��øŲ������� 
              <br>
              <br>
              9�� �����������ǻ��ڻ������Ĺ�����Ϣ����ƽ̨���κ��˲����ҷ���������ҵ���޹ص���Ϣ�����÷���������Ϣ����ɫ��Ϣ�Լ����ַ��ɷ��治����ģ������ɴ˴����ĺ�����ɷ������Ը����������������е��κη������Ρ� 
              <br>
              <br>
              10�� ����������ֻ�ṩ��Ϣ����ƽ̨���������뽻�׹��̣��û������彻�׷�����ʵ��ݣ��Լ����洦���׹����е�ÿһ�����ڣ����ڽ��׹���������һ�о��ף����ɽ���˫������168������Ϣ�����е��κ����Ρ�<br>
              <br>
              11�� �������������û�Ա�ƵĹ���ģʽ���û������ܱ����ṩ�ķ���ǰ�����ȸ�һ�����ã��ʷѱ�׼�����������������ɸ����г������ʱ������ 
              <br>
              <br>
              12�� ��������������Ϣ��ȡ���ŵ���ʽ���οͻ�û�е�½���û������Կ�����Ϣ������Ϣ����ϵ��ʽ�������أ�ֻ����ʽ�����û����ܻ�ȡ�����ķ��� 
              <br>
              <br>
              ѡ����ܱ�ʾͬ�����ϻ�Ա�����������Э�����Ȩ��168������Ϣ���� </p>
            <p></p>
            <p></p>
          </td>
        </tr>
      </table>
      <div align=center> 
        <table cellspacing=0 width=360 border=0>
          <tbody> 
          <tr> 
            <td valign=center align=middle width=198 height=20><a 
            href="User_Reg1.asp" target=_top><img height=21 
            alt=���������һ�� src="images/reg_05.gif" width=56 border=0></a></td>
            <td valign=center align=middle width=197><a 
            href="User_Reg.asp" target=_top><img height=21 
            alt=������һҳ src="images/reg_06.gif" width=56 
        border=0></a></td>
          </tr>
          </tbody> 
        </table>
      </div>
    </td>
  </tr>
  <tr> 
    <td valign=bottom><img height=7 src="images/m_16.gif" 
  width=778></td>
  </tr>
  </tbody> 
</table>
<!--#include file="../inc/down.htm"-->
</BODY></HTML>
