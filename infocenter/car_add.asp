<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../Member/check.asp"-->
<html>
<head>
<title><%=WebName%>-������Դ��Ϣ</title>

<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language=JavaScript src="../inc/p_c_a2.js"></script>
<script LANGUAGE="JavaScript">
function check()
{
if (document.Form1.province.value =="ʡ��"||document.Form1.province2.value=="ʡ��") 
{ 
alert("��ѡ������"); 
document.Form1.province.focus(); 
return (false); 
} 
 
if (document.Form1.carload.value =="") 
{ 
alert("���ز���Ϊ��"); 
document.Form1.carload.focus(); 
return (false); 
}
if (document.Form1.carlong.value =="") 
{ 
alert("��������Ϊ��"); 
document.Form1.carlong.focus(); 
return (false); 
}
if (document.Form1.content.value.length >50) 
{ 
alert('���ܳ���50����!');
document.Form1.content.focus(); 
return (false); 
}

}
</SCRIPT>

<style type="text/css">
<!--
body {
	background-color: #2C68B1;
	margin: 0px;
}
-->
</style>
<link href="../member/images/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
</head>
<BODY>
<table width="750"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="566" align="center" valign="top" bgcolor="#FFFFFF" id="main">
      <table width="98%"  border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="83%" height="38" align="left" id="pos"><a href="http://www.bearingpage.com">����������</a> &gt; ������Դ��Ϣ</td>
          <td width="17%">&nbsp;</td>
      </tr>
    </table>
	  <table width="100%" height="497"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top">
            <table width="98%"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
			  <table width="100%"  border="0">
                             <tr>
                              <td align="left">
                        <div id="TipAll" > <font color="#FF0000">ע���벻Ҫ�����ظ���Ϣ��лл����&nbsp;&nbsp;&nbsp;&nbsp; 
                          </font> </div>                      </td>
                            </tr>
                           </table><%set rs_p=conn.execute("select Count(*) from info2 where GsID='"&session("user")&"'")
						   p_count=rs_p(0)
						   rs_p.close
						   set rs_p=nothing
						   if p_count>=10 then
						      response.write "<p>&nbsp;&nbsp;�Բ�����ѻ�Աֻ�ܷ���10����Ʒ��Ϣ���뷢������Ϣ����ϵ���ǳ�Ϊ��ʽ��Ա��лл��</p>&nbsp;&nbsp;��ѯ�绰��+86-574-63809664 27861860   ���棺+86-574-27860361 ���������� "
						   else
						   %>
  <form method="post" action="car_save.asp" name="Form1" onSubmit="javascript:return check(this);">
		                  
<table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="121" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                       <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">������Դ��Ϣ</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">������:</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600"> 
                          <select id="s1" name="province">
                            <option value="ʡ��">ʡ��</option>
                          </select>
                          <select id="s2" name="city">
                            <option>�ؼ���</option>
                          </select>
                          <select id="s3" name="area">
                            <option>�С��ؼ��С���</option>
                          </select>						  
                          </font></td>
                      </tr>
			<SCRIPT language=javascript>
            setup();
          </SCRIPT>	                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">�����: </td>
                        <td width="81%" valign="top" height="22"><font color="#FF6600">
                          <select id="t1" name="province2">
                            <option value="ʡ��">ʡ��</option>
                          </select>
                          <select id="t2" name="city2">
                            <option>�ؼ���</option>
                          </select>
                          <select id="t3" name="area2">
                            <option>�С��ؼ��С���</option>
                          </select>						  
                          </font> </td>
                      </tr>

		  <SCRIPT language=javascript>
            setup2();
          </SCRIPT>					  
                    				  
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">����:</td>
                        <td width="81%" valign="top" height="22"> 
                        <select name="cartype" id="cartype">
		                   <option value="�г�">�г�</option>
		                   <option value="��ͨ��">��ͨ��</option>
		                   <option value="���">���</option>
		                   <option value="��ҳ�">��ҳ�</option>
		                   <option value="��ʽ��">��ʽ��</option>
		                   <option value="��س�">��س�</option>
		                   <option value="��װ��">��װ��</option>
		                   <option value="ƽ�峵">ƽ�峵</option>
		                   <option value="���س�">���س�</option>
		                   <option value="��ж��">��ж��</option>
		                   <option value="�͹޳�">�͹޳�</option>
		                   <option value="�����">�����</option>
		                   <option value="ǰ�ĺ��">ǰ�ĺ��</option>
		                   <option value="˫�ų�">˫�ų�</option>
		                   <option value="�ӳ���">�ӳ���</option>
		                   <option value="��⳵">��⳵</option>
		                   <option value="������">������</option>
		                   <option value="���³�">���³�</option>
	                </select>						</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">����:</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600">
                          <input id=carload name=carload maxlength="4">
                          ��</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">����:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                          <input id=carlong name=carLong maxlength="4">
                          ��</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF">
                        <td width="19%" height="22">��Ч��:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                        <select name="validate" id="validate">
		                   <option value="5">5��</option>
		                   <option value="15">15��</option>
		                   <option value="30">30��</option>
						   <option value="60">60��</option>
						   <option value="90">90��</option>
						   <option value="120">120��</option>
						   </select>
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">˵����</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF">
						<textarea name=content cols="40" rows="5"></textarea>						</td>
                      </tr>
                      
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="19"> 
                          <input type="submit" value="�� ��"  name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                          <input type="button" onClick="javascript:history.back(1)" value="����">                        </td>
                      </tr>
                    </table>
                  </form> <%end if%>               </td>
            </tr>
          </table>
                        
            </td>
        </tr>
      </table></td>
  </tr>
</table>
</BODY>
</HTML>
