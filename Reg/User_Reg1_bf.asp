<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/config.asp"-->
<%
dim RegUserName,txtVerify
RegUserName=request.form("userid")
session("RegUserName")=RegUserName

if RegUserName="" or strLength(RegUserName)>20 or strLength(RegUserName)<4 then
	founderr=true
	errmsg=errmsg & "�������û���(���ܴ���20С��5)\n"
else
  	if Instr(RegUserName,"<")>0 or Instr(RegUserName,">")>0 or Instr(RegUserName,"=")>0 or Instr(RegUserName,"%")>0 or Instr(RegUserName,chr(32))>0 or Instr(RegUserName,"?")>0 or Instr(RegUserName,"&")>0 or Instr(RegUserName,";")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,"'")>0 or Instr(RegUserName,",")>0 or Instr(RegUserName,chr(34))>0 or Instr(RegUserName,chr(9))>0 or Instr(RegUserName,"��")>0 or Instr(RegUserName,"$")>0 then
		errmsg=errmsg+"�û����к��зǷ��ַ�\n"
		founderr=true
	end if
end if

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
if (document.UserReg.question.value =="") 
{ 
alert("������������ʾ���⣡"); 
document.UserReg.question.focus(); 
return (false); 
} 
if (document.UserReg.answer.value =="") 
{ 
alert("������������ʾ�𰸣�"); 
document.UserReg.answer.focus(); 
return (false); 
} 

if (document.UserReg.qymc.value =="") 
{ 
alert("�����빫˾���ƣ�"); 
document.UserReg.qymc.focus(); 
return (false); 
} 
if (document.UserReg.jyfw.value =="")
{ 
alert("����������ҵ�ľ�Ӫ�ͷ��������"); 
document.UserReg.jyfw.focus(); 
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
var filter2=/^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
if (!filter2.test(document.UserReg.email.value)) { 
alert("�ʼ���ַ����ȷ,��������д��"); 
document.UserReg.email.focus();
document.UserReg.email.select();
return false; 
} 
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
<TABLE cellSpacing=0 cellPadding=0 width=776 align=center border=0>
  <TBODY> 
  <TR> 
    <TD vAlign=top height=323> 
      <TABLE style="BORDER-BOTTOM: #5286b5 1px solid" height=70 cellSpacing=0 
      cellPadding=0 width=776 border=0>
        <TBODY> 
        <TR> 
          <TD width=265 height=64><A href="../"><IMG height=65 
            alt=������ҳ src="images/logo.gif" width=265 align=bottom 
            border=0></A></TD>
          <TD vAlign=bottom align=middle width=511> 
            <TABLE height=50 cellSpacing=0 cellPadding=0 width=489 border=0>
              <TBODY> 
              <TR> 
                <TD width=180><IMG height=50 src="images/title_re1.gif" 
                  width=180></TD>
                <TD width=180><IMG height=50 src="images/title_re2.gif" 
                  width=180></TD>
                <TD width=129><IMG height=50 src="images/title_re3.gif" 
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
            <TABLE cellSpacing=0 cellPadding=10 width="100%" border=0>
              <TBODY> 
              <TR> 
                <TD vAlign=top align=middle width="12%" height=30><IMG 
                  height=37 src="images/e_arrow.gif" width=64></TD>
                <TD align=right width="88%"><IMG height=30 
                  src="images/htxt_01.gif" 
        width=500></TD>
              </TR>
              </TBODY>
            </TABLE>
          </TD>
        </TR>
        </TBODY>
      </TABLE>
      <table border="0" bordercolor="#111111" width="100%" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="100%"> 
            <FORM method="POST" action="User_RegOK.asp" name="UserReg" onSubmit="return FormCheck();" >
			    <input type="hidden" name="userstate" value="new">
                <input type="hidden" name="_fom_grp_" value="Ys"/>
                <input type="hidden" name="M" value="_m_M_"/>
              <table height=51 cellspacing=0 cellpadding=0 width=760 align=center 
      border=0>
                <tbody> 
                <tr> 
                  <td width=257 rowspan=2><img height=51 src="images/dot_101.jpg" 
            width=257></td>
                  <td align=right height=22><span class=S>ע���Ա�����κ������벦��:<span 
            class=style17>0574-63809664</span></span></td>
                </tr>
                <tr> 
                  <td background=images/dot_102.jpg 
      height=29>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <TABLE border=0 bordercolor="#3399FF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-collapse: collapse" cellspacing="0" cellpadding="0" align="center" width="760">
                <tr> 
                  <TD align=right width="200" bgcolor="#f7f7f7"><SPAN 
            class=style16>��Ա�û���:</SPAN></TD>
                  <TD width="13">&nbsp;</TD>
                  <TD width="547"><SPAN class=p14> 
                    <INPUT class=f11 maxLength=20 name="Ys_m_M_user" size="20" value="<%=session("RegUserName")%>" readonly>
                    </SPAN> <font color="#666666"><span class=red>*</span> <span class=huang12>5-20���ַ���<font face="Arial">A-Z</font>�� 
                    <font face="Arial"> a-z</font>�� <font face="Arial"> 0-9</font>����������ȡһ����������ͼ���ĵ�¼��</span></font></TD>
                </tr>
                <TR vAlign="middle" background="images/dot3.gif"> 
                  <TD align=right height="1" colspan="3"></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="0" bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>�� �룺</SPAN></TD>
                  <TD height="13" width="13">&nbsp;</TD>
                  <TD height="13" width="547"> 
                    <INPUT class=f11 maxLength=20 name=pass type=password size="20">
                    <span class=p14><span class=p14> </span></span><span class=red>*</span> 
                    <span class=huang12>��4-20λ����ʹ��Ӣ��(a-z)������(0-9)ע�����ִ�Сд��</span></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="1" colspan="3" background="images/dot3.gif"></TD>
                <TR vAlign="middle"> 
                  <TD align=right height="28" bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>ȷ�����룺</SPAN></TD>
                  <TD height="28" width="13">&nbsp;</TD>
                  <TD height="28" width="547"> 
                    <INPUT class=f11 maxLength=20 name=confirmPassword type=password size="20">
                    <span class=p14><span class=p14> </span></span><span class=red>*</span> 
                    <span class=huang12>���ظ�һ�����������õ�����</span></TD>
                <TR vAlign="middle" > 
                  <TD align=right height="1" colspan="3" background="images/dot3.gif"></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>����������</SPAN></TD>
                  <TD width="13">&nbsp;</TD>
                  <TD height="30" width="547"> 
                    <input class=f11 maxlength=100 name=name size="20">
                    <font color=#990000> </font><span class=red>*</span> <span 
            class=huang12>������ʵ����</span></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="1" colspan="3" background="images/dot3.gif"></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>�Ա�</SPAN></TD>
                  <TD width="13">&nbsp;</TD>
                  <TD height="30" width="547"> 
                    <input class=f11 name=ch type=radio value="����" checked>
                    ����&nbsp;&nbsp; 
                    <input class=f11 name=ch type=radio value="Ůʿ">
                    Ůʿ<span class=red></span></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="1" colspan="3" background="images/dot3.gif"></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>������ʾ���⣺</SPAN></TD>
                  <TD height="49" width="13">&nbsp;</TD>
                  <TD height="49" width="547"> 
                    <select name="question" id="pass_prob">
                      <option selected value="��ѡ��������ʾ����">��ѡ��������ʾ����</option>
                      <option value="�����ĳ����ʲô����">�����ĳ����ʲô����</option>
                      <option value="��ϲ�����˽�ʲô����">��ϲ�����˽�ʲô����</option>
                      <option value="��ϲ��������">��ϲ��������</option>
                      <option value="��ϲ���ĸ���">��ϲ���ĸ���</option>
                      <option value="��ϲ��������">��ϲ��������</option>
                      <option value="��ϲ������">��ϲ������</option>
                      <option value="��ϲ��������">��ϲ��������</option>
                      <option value="��ϲ��������">��ϲ��������</option>
                    </select>
                    <span class=red>*</span> <span 
            class=huang12>����д�Լ�����Ϥ�����⣬���μǣ�������������ʱ���������롣</span></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="1" colspan="3" background="images/dot3.gif"></TD>
                </TR>
                <TR vAlign="middle"> 
                  <TD align=right height="27" bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>������ʾ�𰸣�</SPAN> 
                  <TD height="27" width="13">&nbsp;</TD>
                  <TD height="27" width="547">
                    <INPUT class=f11 maxLength=20 name=answer size="20">
                    <span class=red>*</span> <span 
            class=huang12>���μ�����𰸣��Ա����붪ʧʱ�ش�ϵͳ�����ʡ�������Ļش���ȷ��ϵͳ�ͻ��Զ���������ʾ������</span></TD>
                </TR>
                <INPUT type=hidden 
value=name=which>
                <INPUT name=provincetwo 
type=hidden>
              </TABLE>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td width=257 rowspan=2><img height=53 src="images/dot_201.jpg" 
            width=322></td>
                  <td height=24><img height=1 src="images/dot.gif" width=1></td>
                </tr>
                <tr> 
                  <td background=images/dot_102.jpg 
      height=29>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <TABLE border=0 bordercolor="#3399FF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-collapse: collapse" cellspacing="0" cellpadding="0" align="center" width="760">
                <tr> 
                  <TD align=right height="18" bgcolor="#f7f7f7" width="200"><SPAN 
            class=style16>��˾���ƣ�</SPAN></TD>
                  <TD height="18" width="14">&nbsp;</TD>
                  <TD height="18"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=280><span class=p14><font face=����><span 
                  class=p14><font face=����> 
                          <input style="WIDTH: 214px" 
                  name=qymc>
                          </font></span></font></span></td>
                        <td height=30><span class=red>*</span> <span class=huang12>����д�ڹ��̾�ע��Ĺ�˾/�̺�ȫ��</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="18" bgcolor="#f7f7f7"><SPAN 
            class=style16>��ҵ���ࣺ</SPAN></TD>
                  <TD height="18">&nbsp;</TD>
                  <TD height="30"> 
                    <%
        sql = "select * from Class_1"
        rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	response.write "���������Ŀ��"
	response.end
	else
%>
                    <select name="SortID">
                      <option selected value="<%=trim(rs("SortID"))%>"><%=trim(rs("sort"))%></option>
                      <%      dim selclass
         selclass=rs("SortID")
        rs.movenext
        do while not rs.eof
%>
                      <option value="<%=trim(rs("SortID"))%>"><%=trim(rs("sort"))%></option>
                      <%
        rs.movenext
        loop
	end if
        rs.close
%>
                    </select>
                    <font color=#FF0000 size=2> *</font></TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="18" bgcolor="#f7f7f7"><SPAN 
            class=style16>��Ӫģʽ��</SPAN></TD>
                  <TD height="18">&nbsp;</TD>
                  <TD height="30"><font color="#000080"> 
                    <input type="radio" name="qylb" value="������" checked>
                    ������ 
                    <input type="radio" name="qylb" value="�ɹ���">
                    �ɹ��� 
                    <input type="radio" name="qylb" value="������">
                    ������ <span class=p14><span class=S><span 
            class=red>*</span> </span></span><span 
            class=huang12>ѡ����Ӫ��ҵ</span></font></TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="18" bgcolor="#f7f7f7"><SPAN 
            class=style16>��ҵ���ʣ�</SPAN></TD>
                  <TD height="18">&nbsp;</TD>
                  <TD height="30"> 
                    <select size="1" name="qyxz">
                      <option value="��������">��������</option>
                      <option value="�������">�������</option>
                      <option value="��ҵ��λ">��ҵ��λ</option>
                      <option value="������ҵ">������ҵ</option>
                      <option value="�ɷݼ���">�ɷݼ���</option>
                      <option value="������ҵ">������ҵ</option>
                      <option value="������ҵ">������ҵ</option>
                      <option value="������ҵ">������ҵ</option>
                      <option value="������ҵ">������ҵ</option>
                      <option value="˽Ӫ��ҵ" selected>˽Ӫ��ҵ</option>
                    </select>
                    <font color="#000080">(����)</font> </TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="4" background="images/dot3.gif">
                </tr>
                <tr> 
                  <TD align=right height="18" bgcolor="#f7f7f7"><SPAN 
            class=style16>��Ӫ��Χ��</SPAN> 
                  <TD> 
                  <TD height="18" colspan="2"> 
                    <textarea style="WIDTH: 357px; HEIGHT: 102px" rows="2" name="jyfw" cols="56"></textarea>
                    <font color="#FF0000"> </font><font color="#000080">(����)</font> 
                  </TD>
                </tr>
                <INPUT type=hidden 
value=name=which>
                <INPUT name=provincetwo 
type=hidden>
              </TABLE>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td width=257 rowspan=2><img height=51 src="images/dot_301.jpg" 
            width=260></td>
                  <td height=22><img height=1 src="images/dot.gif" width=1></td>
                </tr>
                <tr> 
                  <td background=images/dot_102.jpg 
      height=29>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <table cellspacing=0 cellpadding=0 width=760 align=center border=0>
                <tbody> 
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                </tbody> 
              </table>
              <TABLE border=0 bordercolor="#3399FF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-collapse: collapse" cellspacing="0" cellpadding="0" align="center" width="760">
                <tr> 
                  <TD align=right bgcolor="#f7f7f7" width="200"> 
                    <p align="right"><span 
            class=style16>��˾���ڵ�:</span> 
                  </TD>
                  <TD height="18" width="14">&nbsp;</TD>
                  <TD height="18"> 
                    <table id=Table2 cellspacing=0 cellpadding=0 width="98%" border=0>
                      <tbody> 
                      <tr> 
                        <td width="56%"> 
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
                        <td class=S width="44%" height="30"><span class=red>*</span> 
                          <span 
                  class=huang12>�㰴�˵�ѡ��</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="27" bgcolor="#f7f7f7"><SPAN 
            class=style16>��˾��ַ��</SPAN></TD>
                  <TD height="27">&nbsp;</TD>
                  <TD height="27"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=250> 
                          <input id=address style="WIDTH: 214px" 
                  name="Ys_m_M_address">
                          &nbsp; </td>
                        <td width=290><span class=red>*</span> <span 
                  class=huang12>ֻ����д�ֵ������ƺ�</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="30" bgcolor="#f7f7f7"><SPAN 
            class=style16>�̶��绰��</SPAN></TD>
                  <TD height="30">&nbsp;</TD>
                  <TD height="30"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=66 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel1 style="WIDTH: 37px" value=+86 name="Ys_m_M_phone1">
                          </span></span>- </td>
                        <td width=99 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel2 style="WIDTH: 53px" name="Ys_m_M_phone2">
                          </span></span>-</td>
                        <td width=82 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel3 style="WIDTH: 103px" name="Ys_m_M_phone3">
                          </span></span></td>
                        <td width=293 height=30><span class=red>*</span> <span 
                  class=huang12>���Ҵ���-����-�绰����</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" colspan="3" background="images/dot3.gif"></TD>
                </tr>
                <tr> 
                  <TD align=right height="30" bgcolor="#f7f7f7"><SPAN 
            class=style16>���棺</SPAN></TD>
                  <TD height="30">&nbsp;</TD>
                  <TD height="30"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=66 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel1 style="WIDTH: 37px" value=+86 name="Ys_m_M_fax1">
                          </span></span>- </td>
                        <td width=99 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel2 style="WIDTH: 53px" name="Ys_m_M_fax2">
                          </span></span>-</td>
                        <td width=82 height=30><span class=p14><span class=p14> 
                          <input 
                  id=tel3 style="WIDTH: 103px" name="Ys_m_M_fax3">
                          </span></span></td>
                        <td width=293 height=30><span 
                class=huang12>�������</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" background="images/dot3.gif" colspan="3"></TD>
                </tr>
                <tr> 
                  <TD align=right height="27" bgcolor="#f7f7f7"><SPAN 
            class=style16>�ֻ���</SPAN></TD>
                  <TD height="27">&nbsp;</TD>
                  <TD height="30"> 
                    <input id=handset style="WIDTH: 214px" 
            name=mobile>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" background="images/dot3.gif" colspan="3"></TD>
                </tr>
                <tr> 
                  <TD align=right height="29" bgcolor="#f7f7f7"><SPAN 
            class=style16>�ʱࣺ</SPAN></TD>
                  <TD height="29">&nbsp;</TD>
                  <TD height="30"> 
                    <input id=post style="WIDTH: 214px" 
            name=post>
                    &nbsp;<span class=huang12>�ʱ����</span> </TD>
                </tr>
                <tr> 
                  <TD height="1" background="images/dot3.gif" colspan="3"></TD>
                </tr>
                <tr> 
                  <TD align=right height="29" bgcolor="#f7f7f7"><SPAN 
            class=style16>E-mail��</SPAN></TD>
                  <TD height="29">&nbsp;</TD>
                  <TD height="29"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=280><span class=p14> 
                          <input style="WIDTH: 214px" name=email>
                          </span></td>
                        <td height=30><span class=red>*</span> <span 
                  class=huang12>���ĵ��������ַ,����&nbsp; info@cx56w.com</span></td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" background="images/dot3.gif" colspan="3"> 
                </tr>
                <tr> 
                  <TD align=right height="29" bgcolor="#f7f7f7"><SPAN 
            class=style16>��ַ��</SPAN></TD>
                  <TD height="29">&nbsp;</TD>
                  <TD height="29"> 
                    <table cellspacing=0 cellpadding=0 width=540 border=0>
                      <tbody> 
                      <tr> 
                        <td width=247><span class=p14> 
                          <input id=url 
                  style="WIDTH: 214px" name=web value="http://">
                          </span></td>
                        <td class=huang12 width=293 
              height=30>������ԭ����վ��ַ���ڳ����������õ�һ���ƹ�</td>
                      </tr>
                      </tbody> 
                    </table>
                  </TD>
                </tr>
                <tr> 
                  <TD height="1" background="images/dot3.gif" colspan="3"></TD>
                </tr>
                <tr> 
                  <TD align=right height="29" bgcolor="#f7f7f7"><SPAN 
            class=style16>��˾���:</SPAN></TD>
                  <TD height="29">&nbsp;</TD>
                  <TD height="29"> 
                    <textarea style="WIDTH: 357px; HEIGHT: 102px" rows="5" name="qyjj" cols="56"></textarea>
                    <span 
            class=red>*</span> <span class=huang12>��˾���������������25��</span></TD>
                </tr>
                <tr>
                  <TD align=right height="29" bgcolor="#f7f7f7">��֤�룺</TD>
                  <TD height="29">&nbsp;</TD>
                  <TD height="29">
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
&nbsp;��������֤����ˢ�� </TD>
                </tr>
              </TABLE>
              <table border="0" cellpadding="0" style="border-collapse: collapse" width="768" id="AutoNumber6" height="88" cellspacing="0">
                <tr> 
                  <td width="100%" height="40" align="center">�� <a class=red 
            href="Item.asp" 
            target=_blank>����Ķ�������������������</a> ��</td>
                </tr>
                <tr> 
                  <td width="100%"> 
                    <p align="center"> 
                      <input class=button04 type="submit" value="���Ѿ��Ķ���ͬ����ܷ�������,�����ҵĳ�����������Ա�ʺ�" >
                    </p>
                  </td>
                </tr>
              </table>
            </FORM>
          </td>
        </tr>
      </table>
    </TD>
  </TR>
  </TBODY>
</TABLE>
<!--#include file="../inc/down.htm"-->
</body>
</html>
<%end sub%>