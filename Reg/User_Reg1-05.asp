<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/config.asp"-->
<%
dim RegUserName,txtVerify
RegUserName=request.form("userid")
session("RegUserName")=RegUserName

'if RegUserName="" or strLength(RegUserName)>20 or strLength(RegUserName)<4 then
	'founderr=true
	'errmsg=errmsg & "�������û���(���ܴ���20С��5)\n"
'else
  	'if Instr(RegUserName,"<")>0 or Instr(RegUserName,">")>0 or Instr(RegUserName,"=")>0 or Instr(RegUserName,"%")>0 or Instr(RegUserName,chr(32))>0 or Instr(RegUserName,"?")>0 or Instr(RegUserName,"&")>0 or Instr(RegUserName,";")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,"'")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,chr(34))>0 or Instr(RegUserName,chr(9))>0 or Instr(RegUserName,"��")>0 or Instr(RegUserName,"$")>0 then
		'errmsg=errmsg+"�û����к��зǷ��ַ�\n"
		'founderr=true
	'end if
'end if

if founderr=false then
	dim sqlReg,rsReg,sqladminname,rsadminname

	sqladminname="select * from Manage_user where UserName='" & RegUserName & "'"
	set rsadminname=conn.execute(sqladminname)
	
	if not(rsadminname.bof and rsadminname.eof) then
		founderr=true
		errmsg=errmsg & "���ѣ���ð�����Ա����һ���û������ϴ�\n"
		session("RegUserName")=""
	else
	
    set rs=server.CreateObject ("adodb.recordset")
      sql="select * from qyml where user='"&RegUserName&"'"
      rs.open sql,conn,1,3
      if not (rs.eof and rs.bof) then
           founderr=true
		   errmsg=errmsg & "��ע����û��Ѿ����ڣ��뻻һ���û��������ԣ�\n"
		   session("RegUserName")=""
      end if
	  rs.close
	  set rs=nothing
	end if
end if

if founderr=false then
	call RegNext()
else
	call WriteRegErrmsg()
end if 

Sub RegNext()
%>
<html>
<head>
<title>��Աע�ᡪ������������</title>
<%
dim rs
dim sql
set rs=server.createobject("adodb.recordset")
%>


<script Language="JavaScript"> 
function FormCheck()
{ 
if (document.UserReg.pass.value =="") 
{
alert("����д�������룡");
document.UserReg.pass.focus();
return false; 
}
if(document.UserReg.confirmPassword.value==""){
alert("����������ȷ�����룡");
document.UserReg.confirmPassword.focus();
return false;
}
var filter=/^\s*[.A-Za-z0-9_-]{5,15}\s*$/;
if (!filter.test(document.UserReg.pass.value)) { 
alert("������д����ȷ,��������д����ʹ�õ��ַ�Ϊ��A-Z a-z 0-9 _ - .)���Ȳ�С��5���ַ���������15���ַ���ע�ⲻҪʹ�ÿո�"); 
document.UserReg.pass.focus();
document.UserReg.pass.select();
return false; 
} 
if (document.UserReg.pass.value!=document.UserReg.confirmPassword.value ){
alert("������д�����벻һ�£���������д��"); 
document.UserReg.pass.focus();
document.UserReg.pass.select();
return false; 
} 
if (document.UserReg.name.value =="") 
{ 
alert("����������������"); 
document.UserReg.name.focus(); 
return (false); 
} 

if (document.UserReg.qymc.value =="") 
{ 
alert("�����빫˾���ƣ�"); 
document.UserReg.qymc.focus(); 
return (false); 
} 

if (document.UserReg.siji.value =="") 
{ 
alert("������˾�����ƣ�"); 
document.UserReg.siji.focus(); 
return (false); 
} 

if (document.UserReg.Ys_m_M_province.value =="ʡ��") 
{ 
alert("��ѡ������"); 
document.UserReg.Ys_m_M_province.focus(); 
return (false); 
} 
if (document.UserReg.Ys_m_M_city.value =="") 
{ 
alert("��ѡ���������"); 
document.UserReg.Ys_m_M_city.focus(); 
return (false); 
} 
if (document.UserReg.Ys_m_M_area.value =="") 
{ 
alert("��ѡ���������"); 
document.UserReg.Ys_m_M_area.focus(); 
return (false); 
}
if (document.UserReg.Ys_m_M_address.value =="") 
{ 
alert("������������ϵ��ַ��"); 
document.UserReg.Ys_m_M_address.focus(); 
return (false); 
} 

if (document.UserReg.Ys_m_M_phone2.value ==""||document.UserReg.Ys_m_M_phone3.value =="") 
{ 
alert("������������ϵ�绰��"); 
document.UserReg.Ys_m_M_phone2.focus(); 
return (false); 
} 

var regexp = /^[0123456789]{1,6}$/g; 
if (regexp.test(document.UserReg.post.value)!=1)
{
alert("��������ȷ���������룡");
document.UserReg.post.focus(); 
document.UserReg.post.select();
return (false);
}
//var filter2=/^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
//if (!filter2.test(document.UserReg.email.value)) { 
//alert("�ʼ���ַ����ȷ,��������д��"); 
//document.UserReg.email.focus();
//document.UserReg.email.select();
//return false; 
//} 
if (document.UserReg.qyjj.value =="")
{ 
alert("�����빫˾��飡"); 
document.UserReg.qyjj.focus(); 
return (false); 
} 
if( document.UserReg.txtVerify.value =="") {
		alert("\������������֤�� !");
		document.UserReg.txtVerify.focus();
		return false;
}
if( document.UserReg.txtVerify.value.length!=4 || fIsNumber(document.UserReg.txtVerify.value,"1234567890")!=1) {
		alert("\��֤��ֻ��Ϊ�ұ�ͼƬ��ʾ4λ����!");
		document.UserReg.txtVerify.focus();
		return false;
	}
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
<script language=JavaScript src="../inc/p_c_a.js"></script>
<head>
<LINK href="images/css.css" type="text/css" rel="stylesheet">
<style type="text/css">.p14 { FONT-SIZE: 14px; COLOR: #000000; LINE-HEIGHT: 130% }
	.S { FONT-SIZE: 12px }
	.style16 { FONT-WEIGHT: bold }
	.button { BORDER-RIGHT: #ffffff thin; BACKGROUND-POSITION: center center; BORDER-TOP: #ffffff thin; FONT-SIZE: 12px; BORDER-LEFT: #ffffff thin; WIDTH: 50px; COLOR: #ffffff; BORDER-BOTTOM: #ffffff thin; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 20px; BACKGROUND-COLOR: #999999 }
	.button01 { BORDER-RIGHT: #9c9a9c 1px solid; BACKGROUND-POSITION: center center; BORDER-TOP: #9c9a9c 1px solid; FONT-SIZE: 12px; MARGIN: 1px; BORDER-LEFT: #9c9a9c 1px solid; WIDTH: 100px; COLOR: #666666; BORDER-BOTTOM: #9c9a9c 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: auto; BACKGROUND-COLOR: #cccccc }
	.button02 { BORDER-RIGHT: #9c9a9c 1px solid; BACKGROUND-POSITION: center center; BORDER-TOP: #9c9a9c 1px solid; FONT-SIZE: 12px; MARGIN: 1px; BORDER-LEFT: #9c9a9c 1px solid; WIDTH: 50px; COLOR: #666666; BORDER-BOTTOM: #9c9a9c 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 20px; BACKGROUND-COLOR: #cccccc }
	.button04 { BORDER-RIGHT: #aaaaaa 1px solid; BORDER-TOP: #fefefe 1px solid; FONT-SIZE: 12px; BACKGROUND: #ececec; BORDER-LEFT: #fefefe 1px solid; WIDTH: 380px; COLOR: #000000; PADDING-TOP: 1px; BORDER-BOTTOM: #aaaaaa 1px solid; WHITE-SPACE: normal; HEIGHT: 23px }
	.button03 { BORDER-RIGHT: #aaaaaa 1px solid; BORDER-TOP: #fefefe 1px solid; FONT-SIZE: 12px; BACKGROUND: #ececec; BORDER-LEFT: #fefefe 1px solid; COLOR: #000000; PADDING-TOP: 1px; BORDER-BOTTOM: #aaaaaa 1px solid; WHITE-SPACE: normal; LETTER-SPACING: 6px; HEIGHT: 23px }
	.style17 { FONT-WEIGHT: bold; COLOR: #cc3333 }
	</style>
</head>

<BODY onLoad="setup()">
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
<table width="776" border="0" align="center" cellspacing="0" cellpadding="0" class=huang12>
  <tr> 
    <td width="231" valign="top" style="border-collapse: collapse; LINE-HEIGHT: 150%" bgcolor="#FFFAE1"> 
      <p align=center><b><font size=4>�����д��Ա���� </font></b></p>
      <ol>
        <li>��<font color=#ff0000> * </font>����ʾ����� 
        <li>���ڵ������뾫ȷ�����ؼ�(�������������ѡ)�� 
        <li>�Ƽ���Ա����Ǳ�Ļ�Ա�Ƽ���������Ǹ���Ա�ı��,���û�����Ƽ�,���Բ���. 
        <li>����д�ĵ�λ���ܽ�������������ҳ�ϣ������д������������������ҳ�ϵĵ�λ����һ�����ǿհף����Ӱ��������ҵ���� 
        <li>������е����ʼ������ַ��������ϣ���������˵�¼���룬ϵͳ�Ὣ���뷢������������,Ҳ�Ա����ǻ�������������ʱ��ͨ�� 
        <li>ע��ɹ���ϵͳ��Ϊ������һ����Ա��ţ����ǻ�Ա��Ψһ��ʶ�����μǣ� 
        <li>ע��ɹ������Ϳ��Ե�¼��վ�ˣ������ѡ���ˣ�����Զ���¼�������Ժ��ٷ������ǵ���վʱ�����Զ���¼������ÿ�ζ������Ա��ź����룮 
        <li><font 
        color=#ff0066>ע�⣺��ע��ɹ����������վ��ҳ���������������ĵ�λ���ƻ����ˣ��Ⲣ��������ע�᲻�ɹ�����Ϊ��ҳ��һ���ӳ٣����5���ӣ����벻Ҫ�ٴ��ظ�ע�ᣡ</font> 
        </li>
      </ol>
    </td>
    <td width="545"> 
      <table border="0" bordercolor="#111111" cellspacing="0" cellpadding="0" align="center" width="99%">
        <tr> 
    <td> 
      <form method="POST" action="User_RegOK.asp" name="UserReg" onSubmit="return FormCheck();" >
        <input type="hidden" name="userstate" value="new">
        <input type="hidden" name="_fom_grp_" value="Ys"/>
        <input type="hidden" name="M" value="_m_M_"/>
        <table height=25 cellspacing=0 cellpadding=4 align=center 
      border=0 width="100%">
          <tbody> 
          <tr> 
            <td bgcolor="#CCCCCC">�û���������<span class=S></span></td>
          </tr>
          </tbody> 
        </table>
        <table border=0 bordercolor="#3399FF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-collapse: collapse" cellspacing="0" cellpadding="0" align="center" width="100%" >
          <tr> 
                  <td align=right  bgcolor="#f7f7f7" width="101"><span 
            class=style16>�� �� ��:</span></td>
                  <td width="9">&nbsp;</td>
                  <td width="430"><span class=p14> 
                    <input class=f11 maxlength=20 name="Ys_m_M_user" size="20" value="">
                    </span> <font color="#666666"><span class=red>*</span> <span class=huang12>5-20���ַ���<font face="Arial">A-Z</font>�� 
                    <font face="Arial"> a-z</font>�� <font face="Arial"> 0-9</font>��</span></font></td>
          </tr>
          <tr valign="middle" background="images/dot3.gif"> 
            <td align=right height="1" colspan="3"></td>
          </tr>
          <tr valign="middle"> 
                  <td align=right height="0" bgcolor="#f7f7f7" width="101" ><span 
            class=style16>�ܡ�����:</span></td>
                  <td height="13" width="9">&nbsp;</td>
                  <td height="13" width="430"> 
                    <input class=f11 maxlength=20 name=pass type=password size="20">
                    <span class=p14><span class=p14> </span></span><span class=red>*</span> 
                    <span class=huang12>4-20λ������Ӣ��(a-z)������(0-9)ע�����ִ�Сд</span></td>
          </tr>
          <tr valign="middle"> 
            <td align=right height="1" colspan="3" background="images/dot3.gif"></td>
          <tr valign="middle"> 
                  <td align=right height="28" bgcolor="#f7f7f7" width="101" ><span 
            class=style16>ȷ������:</span></td>
                  <td height="28" width="9">&nbsp;</td>
                  <td height="28" width="430"> 
                    <input class=f11 maxlength=20 name=confirmPassword type=password size="20">
              <span class=p14><span class=p14> </span></span><span class=red>*</span> 
              <span class=huang12>���ظ�һ�����������õ�����</span></td>
          <tr valign="middle" > 
            <td align=right height="1" colspan="3" background="images/dot3.gif"></td>
          </tr>
          <tr valign="middle"> 
                  <td align=right bgcolor="#f7f7f7" width="101" ><span 
            class=style16>��������:</span></td>
                  <td width="9">&nbsp;</td>
                  <td height="30" width="430"> 
                    <input class=f11 maxlength=100 name=name size="20">
              <font color=#990000> </font><span class=red>*</span> <span 
            class=huang12>������ʵ����</span></td>
          </tr>
          <tr valign="middle"> 
            <td align=right height="1" colspan="3" background="images/dot3.gif"></td>
          </tr>
          <tr valign="middle"> 
                  <td align=right bgcolor="#f7f7f7" width="101" ><span 
            class=style16>�ԡ�����:</span></td>
                  <td width="9">&nbsp;</td>
                  <td height="30" width="430"> 
                    <input class=f11 name=ch type=radio value="����" checked>
              ����&nbsp;&nbsp; 
              <input class=f11 name=ch type=radio value="Ůʿ">
              Ůʿ<span class=red></span></td>
          </tr>
          <tr valign="middle"> 
            <td align=right height="1" colspan="3" background="images/dot3.gif"></td>
          </tr>
          <tr valign="middle"> 
            <td align=right height="1" colspan="3" background="images/dot3.gif"></td>
          </tr>
          <input type=hidden 
value=name=which name="hidden">
          <input name=provincetwo 
type=hidden>
        </table>
        <table cellspacing=0 cellpadding=0 align=center border=0>
          <tbody> 
          <tr> 
            <td>&nbsp;</td>
          </tr>
          </tbody> 
        </table>
        <table cellspacing=0 cellpadding=4 align=center border=0 width="100%" height="25">
          <tbody> 
          <tr> 
            <td bgcolor="#CCCCCC">��˾���Ƽ���Ӫҵ��<img height=1 src="images/dot.gif" with=1></td>
          </tr>
          </tbody> 
        </table>
              <table border=0 bordercolor="#3399FF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-collapse: collapse" cellspacing="0" cellpadding="0" align="center" width="100%" >
                <tr> 
                  <td align=right height="18" bgcolor="#f7f7f7" width="101" ><span 
            class=style16>��˾����:</span></td>
                  <td height="18" with="14" width="9">&nbsp;</td>
                  <td height="18" width="430"> 
                    <table cellspacing=0 cellpadding=0  border=0>
                      <tbody> 
                      <tr> 
                        <td><span class=p14><font face=����><span 
                  class=p14><font face=����> 
                          <input  
                  name=qymc>
                          </font></span></font></span></td>
                        <td height=30><span class=red>*</span> <span class=huang12>����д�ڹ��̾�ע��Ĺ�˾/�̺�ȫ��</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </td>
          </tr>
          <tr> 
            <td height="1" colspan="3" background="images/dot3.gif"></td>
          </tr>
          <tr> 
                  <td align=right height="18" bgcolor="#f7f7f7" width="101"><span 
            class=style16>��ҵ����:</span></td>
                  <td height="18" width="9">&nbsp;</td>
                  <td height="30" width="430"> 
                    <%
        sql = "select * from Class_1"
        rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	response.write "���������Ŀ��"
	response.end
	else
%>
                    <select name="SortID">
                <option selected value="<%=trim(rs("sort"))%>"><%=trim(rs("sort"))%></option>
                <%      dim selclass
         selclass=rs("SortID")
        rs.movenext
        do while not rs.eof
%>
                <option value="<%=trim(rs("sort"))%>"><%=trim(rs("sort"))%></option>
                <%
        rs.movenext
        loop
	end if
        rs.close
%>
              </select>
                    <font color=#FF0000>*</font></td>
          </tr>
          <tr> 
            <td height="1" colspan="3" background="images/dot3.gif"></td>
          </tr>
          <tr> 
            <td height="1" colspan="3" background="images/dot3.gif"></td>
          </tr>
          <tr> 
                  <td align=right height="18" bgcolor="#f7f7f7" class="style16" width="101">��Ҫ˾��:</td>
                  <td height="18" width="9">&nbsp;</td>
                  <td height="18" width="430"> 
                    <input type="text" name="siji">
                    <font color="#FF0000">*</font></td>
          </tr>
          <input type=hidden 
value=name=which name="hidden">
          <input name=provincetwo 
type=hidden>
        </table>
        <table cellspacing=0 cellpadding=0 align=center border=0>
          <tbody> 
          <tr> 
            <td>&nbsp;</td>
          </tr>
          </tbody> 
        </table>
        <table cellspacing=0 cellpadding=4 align=center border=0 width="100%">
          <tbody> 
          <tr> 
            <td bgcolor="#CCCCCC">��˾��ϵ��ʽ<img height=1 src="images/dot.gif" with=1></td>
          </tr>
          </tbody> 
        </table>
        <table border=0 bordercolor="#3399FF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-collapse: collapse" cellspacing="0" cellpadding="0" align="center" >
          <tr> 
                  <td align=right bgcolor="#f7f7f7" width="101" > 
                    <p align="right"><span 
            class=style16>��˾���ڵ�:</span>
                  </td>
                  <td height="18" with="14" width="9">&nbsp;</td>
                  <td height="18" width="430"> 
                    <table id=Table2 cellspacing=0 cellpadding=0 with="98%" border=0>
                <tbody> 
                <tr> 
                  <td with="56%"> 
                    <select id="s1" name="Ys_m_M_province">
                      <option value="ʡ��">ʡ��</option>
                    </select>
                    <select id="s2" name="Ys_m_M_city">
                      <option>�ؼ���</option>
                    </select>
                    <select id="s3" name="Ys_m_M_area">
                      <option>�С��ؼ��С���</option>
                    </select>
                  </td>
                  <td class=S with="44%" height="30"><span class=red>*</span> 
                    <span 
                  class=huang12>�㰴�˵�ѡ��</span></td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr> 
            <td height="1" colspan="3" background="images/dot3.gif"></td>
          </tr>
          <tr> 
                  <td align=right height="27" bgcolor="#f7f7f7" width="101"><span 
            class=style16>��˾��ַ:</span></td>
                  <td height="27" width="9">&nbsp;</td>
                  <td height="27" width="430"> 
                    <table cellspacing=0 cellpadding=0  border=0>
                <tbody> 
                <tr> 
                  <td with=250> 
                    <input id=address  
                  name="Ys_m_M_address">
                    &nbsp; </td>
                  <td with=290><span class=red>*</span> <span 
                  class=huang12>ֻ����д�ֵ������ƺ�</span></td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr> 
            <td height="1" colspan="3" background="images/dot3.gif"></td>
          </tr>
          <tr> 
                  <td align=right height="30" bgcolor="#f7f7f7" width="101"><span 
            class=style16>�̶��绰:</span></td>
                  <td height="30" width="9">&nbsp;</td>
                  <td height="30" width="430"> 
                    <table cellspacing=0 cellpadding=0  border=0>
                <tbody> 
                <tr> 
                  <td with=17 height=30><span class=p14><span class=p14> 
                    <input 
                  id=tel1 style="with: 37px" value=+86 name="Ys_m_M_phone1" type="hidden">
                    </span></span></td>
                  <td with=75 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel2 style="with: 53px" name="Ys_m_M_phone2" size="4">
                    </span></span>-</td>
                  <td with=170 height=30><span class=p14><span class=p14> 
                    <input 
                  id=tel3 style="with: 103px" name="Ys_m_M_phone3">
                    </span></span></td>
                  <td with=278 height=30><span class=red>*</span> <span 
                  class=huang12>���Ҵ���-����-�绰����</span></td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr> 
            <td height="1" colspan="3" background="images/dot3.gif"></td>
          </tr>
          <tr> 
                  <td align=right height="30" bgcolor="#f7f7f7" width="101"><span 
            class=style16>��������:</span></td>
                  <td height="30" width="9">&nbsp;</td>
                  <td height="30" width="430"> 
                    <table cellspacing=0 cellpadding=0  border=0>
                <tbody> 
                <tr> 
                  <td with=17 height=30><span class=p14><span class=p14> 
                    <input 
                  id=tel1 style="with: 37px" value=+86 name="Ys_m_M_fax1" type="hidden">
                    </span></span></td>
                  <td with=78 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel2 style="with: 53px" name="Ys_m_M_fax2" size="4">
                    </span></span>-</td>
                  <td with=167 height=30><span class=p14><span class=p14> 
                    <input 
                  id=tel3 style="with: 103px" name="Ys_m_M_fax3">
                    </span></span></td>
                  <td with=278 height=30><span 
                class=huang12>�������</span></td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr> 
            <td height="1" background="images/dot3.gif" colspan="3"></td>
          </tr>
          <tr> 
                  <td align=right height="27" bgcolor="#f7f7f7" width="101"><span 
            class=style16>�֡�����:</span></td>
                  <td height="27" width="9">&nbsp;</td>
                  <td height="30" width="430"> 
                    <input id=handset  name=mobile>
            </td>
          </tr>
          <tr> 
            <td height="1" background="images/dot3.gif" colspan="3"></td>
          </tr>
          <tr> 
                  <td align=right height="29" bgcolor="#f7f7f7" width="101"><span 
            class=style16>�ʡ�����:</span></td>
                  <td height="29" width="9">&nbsp;</td>
                  <td height="30" width="430"> 
                    <input id=post  
            name=post>
              &nbsp;<span class=huang12>�ʱ����</span> </td>
          </tr>
          <tr> 
            <td height="1" background="images/dot3.gif" colspan="3"></td>
          </tr>
          <tr> 
                  <td align=right height="29" bgcolor="#f7f7f7" width="101"><span 
            class=style16>E-mail:</span></td>
                  <td height="29" width="9">&nbsp;</td>
                  <td height="29" width="430"> 
                    <table cellspacing=0 cellpadding=0  border=0>
                <tbody> 
                <tr> 
                  <td><span class=p14> 
                    <input  name=email>
                    </span></td>
                  <td height=30><span 
                  class=huang12>���ĵ��������ַ,����&nbsp; info@cx56w.com</span></td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr> 
            <td height="1" background="images/dot3.gif" colspan="3"> 
          </tr>
          <tr> 
                  <td align=right height="29" bgcolor="#f7f7f7" width="101"><span 
            class=style16>������ַ:</span></td>
                  <td height="29" width="9">&nbsp;</td>
                  <td height="29" width="430"> 
                    <table cellspacing=0 cellpadding=0  border=0>
                <tbody> 
                <tr> 
                  <td with=247><span class=p14> 
                    <input id=url 
                   name=web value="http://">
                    </span></td>
                  <td class=huang12  
              height=30>������ԭ����վ��ַ���ڳ����������õ�һ���ƹ�</td>
                </tr>
                </tbody> 
              </table>
            </td>
          </tr>
          <tr> 
            <td height="1" background="images/dot3.gif" colspan="3"></td>
          </tr>
          <tr> 
                  <td align=right height="29" bgcolor="#f7f7f7" width="101"><span 
            class=style16>��˾���:</span></td>
                  <td height="29" width="9">&nbsp;</td>
                  <td height="29" width="430"> 
                    <textarea style="with: 357px; HEIGHT: 102px" rows="5" name="qyjj" cols="56"></textarea>
              <span 
            class=red>*</span> <span class=huang12>��˾���������������25��</span></td>
          </tr>
          <tr> 
                  <td align=right height="29" bgcolor="#f7f7f7" width="101">�� 
                    ֤ ��:</td>
                  <td height="29" width="9">&nbsp;</td>
                  <td height="29" width="430"> 
                    <input id=txtVerify name=txtVerify  type="text" size="4">
              <script>
                    document.write("<img id=\"vcodeimage\" src=\"GetCode.asp\"  alt=\"��֤��,�������?����ˢ����֤��\" style=\"cursor : pointer;\" onclick=\"RefreshVerifyCode();\" >")

                    var obImg = document.getElementById("vcodeimage");
                    var timeoutid=null; 

                    function RefreshVerifyCode()
                    {
                      clearTimeout(timeoutid);
                      obImg.src = "GetCode.asp"
                      timeoutid = setTimeout("RefreshVerifyCode()", 240000);
                    }

                    timeoutid = setTimeout("RefreshVerifyCode()", 240000);
                  </script>
              &nbsp;��������֤����ˢ�� </td>
          </tr>
        </table>
              <table border="0" cellpadding="0" style="border-collapse: collapse" with="768" id="AutoNumber6" height="40" cellspacing="0" align="center">
                <tr> 
                  <td> 
                    <p align="center"> 
                      <input class=button04 type="submit" value="���Ѿ��Ķ���ͬ����ܷ�������,�����ҵĳ�����������Ա�ʺ�" name="submit" >
                    </p>
                  </td>
                </tr>
              </table>
      </form>
    </td>
  </tr>
</table>
</td>
  </tr>
</table>
<!--#include file="../inc/down.htm"-->
</body>
</html>
<%end sub%>