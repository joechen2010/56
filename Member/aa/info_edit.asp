<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<title>�޸Ĺ�����Ϣ</title>
<script LANGUAGE="JavaScript">
function check()
{
if (document.Form1.showname.value=="")
{
alert("��Ϣ���ⲻ��Ϊ��")
document.Form1.showname.focus()
document.Form1.showname.select()
return
}
if (document.Form1.type.value =="") 
{ 
alert("��ѡ�������"); 
document.Form1.type.focus(); 
return (false); 
} 
if (document.Form1.content.value=="")
{
alert("��Ϣ���ݲ���Ϊ��")
document.Form1.content.focus()
document.Form1.content.select() 
return
}
if (document.Form1.linkman.value=="")
{
alert("��ϵ�˲���Ϊ��")
document.Form1.linkman.focus()
document.Form1.linkman.select() 
return
}
if (document.Form1.phone.value=="")
{
alert("��ϵ�绰����Ϊ��")
document.Form1.phone.focus()
document.Form1.phone.select() 
return
}
if (document.Form1.address.value=="")
{
alert("��ϵ��ַ����Ϊ��")
document.Form1.address.focus()
document.Form1.address.select() 
return
}
if (document.Form1.country.value=="")
{
alert("����������Ϊ��")
document.Form1.country.focus()
document.Form1.country.select() 
return
}
if (document.Form1.city.value=="")
{
alert("���ڳ��в���Ϊ��")
document.Form1.city.focus()
document.Form1.city.select() 
return
} 
if (document.Form1.email.value=="")
{
alert("�����ʼ�����Ϊ��")
document.Form1.email.focus()
document.Form1.email.select() 
return
} 
if (document.Form1.postcode.value=="")
{
alert("�������벻��Ϊ��")
document.Form1.postcode.focus()
document.Form1.postcode.select() 
return
} 
document.Form1.submit()
}
</SCRIPT>
<%
if not isEmpty(request("info_id")) then
info_id=request("info_id")
else
inf_id=1
end if
sql="select * from info where info_id="+cstr(info_id)+" order by info_ID desc"
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
%>
<script language="javascript">
function show_sader(mylink)
{
window.open(mylink,'','top=50,left=50,width=430,height=400,scrollbars=no')
}
</script>
<style type="text/css">
<!--
body {
	background-color: #2C68B1;
	margin: 0px;
}
-->
</style>
<link href="images/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
</head>
<BODY>
<!--#include file="head.htm"-->
<table width="750"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">����������</a> &gt; �޸Ĺ�����Ϣ</td>
          <td width="17%"><img src="images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="534"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> <%=session("user")%></font> </strong>����ӭ��������</td>
        </tr>
      </table>
      <table width="100%" height="497"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top"><table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="images/title_2.jpg" bgcolor="#FFFFFF" >
                        <div align="center"><font color="#FFFFFF">�޸Ĺ�����Ϣ </font></div>
                      </td>
                      <td width="13" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_3.jpg" width="15" height="25"></td>
                      <td width="392" valign="middle">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
			  <table width="100%"  border="0">
                             <tr>
                              <td align="left"><div id="TipAll" >
                               ���������б���*����Ŀ��
                              </div></td>
                            </tr>
                           </table>
                  <table border="0" cellspacing="1" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="154" bgcolor="#B1D4F2">
                    <form method="POST" action="info_Edit_save.asp?info_id=<%=rs("info_id")%>" name="Form1">
                      <tr> 
                        <td colspan="2" height="19" background="images/title_bg.gif"><b><font color="#FFFFFF"><span style="font-size: 10.5pt">Ҫ�޸ĵ���Ϣ</span></font></b></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF">��Ϣ���⣺</td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> <font color="#FF6600"> 
                          *
                          <input type="text" size="21" name="showname" style="border-style: solid; border-width: 1" value="<%=rs("showname")%>">
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF">�������</td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <select size="1" style="font-size: 9pt" name="type">
                            <option value="<%=rs("type")%>" selected><%=rs("type")%></option>
                            <option value="�ɹ�">�ɹ�</option>
                            <option value="��Ӧ">��Ӧ</option>
                            <option value="����">����</option>
                          </select>
                          <font color="#FF6600"> *</font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF">�� Ч �ڣ�</td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <select size="1" name="period" style="font-size: 9pt">
                            <option selected value="<%=rs("period")%>"><%=rs("period")%></option>
                            <option value="һ��">һ��</option>
                            <option value="����">����</option>
                            <option value="һ����">һ����</option>
                            <option value="������">������</option>
                            <option value="������">������</option>
                            <option value="һ��">һ��</option>
                            <option value="����">����</option>
                          </select>
                          <font color="#FF6600"> *</font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF">��Ϣ���ݣ�<font color="#800000"><br>
                          ������ϸ�����������Ĺ�����Ϣ��</font></td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <textarea rows="8" cols="33" name="content" class="f11" style="border-style:solid; border-width:1; font-family: ����"><%=dvHTMLCode(rs("content"))%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td colspan="2" height="19" background="images/title_bg.gif"><font color="#FFFFFF" style="font-size: 10.5pt"><b>������ϵ����</b></font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF"> �� ϵ �ˣ� 
                        </td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <input type="text" class="smallInput" size="21" name="linkman" value=<%=rs("linkman")%>>
                          <font color="#FF6600"> *</font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF">��ϵ�绰��</td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <input type="text" class="smallInput" size="21" name="phone" value=<%=rs("phone")%>>
                          <font color="#FF6600"> *</font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF"> ��˾���ƣ� </td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <input type="text" class=smallInput size="21" name="company" value=<%=rs("company")%>>
                          <font color="#FF6600"> </font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF">��˾��ַ��</td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <input class=smallInput type="text" size="21" name="address" value=<%=rs("address")%>>
                          <font color="#FF6600"> *</font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF"> �������� </td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <input type="text" class=smallInput size="21" name="country" value=<%=rs("country")%>>
                          <font color="#FF6600"> *</font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF">��������</td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <input type="text" class=smallInput size="21" name="city" value=<%=rs("city")%>>
                          <font color="#FF6600"> *</font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF"> �����ʼ��� </td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <input class="smallInput" type="text" size="21" name="email" value=<%=rs("mail")%>>
                          <font color="#FF6600"> *</font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="19" bgcolor="#EDF6FF">�������룺</td>
                        <td width="34%" height="19" bgcolor="#EDF6FF"> 
                          <input class="smallInput" type="text" size="16" name="postcode" value=<%=rs("postcode")%>>
                          <font color="#FF6600"> *</font></td>
                      </tr>
                      <tr> 
                        <td width="15%" height="24" bgcolor="#EDF6FF"> ��˾���棺 </td>
                        <td width="34%" height="24" bgcolor="#EDF6FF"> 
                          <input class="smallInput" type="text" size="21" name="fax" value=<%=rs("fax")%>>
                        </td>
                      </tr>
                      <tr> 
                        <td width="15%" height="24" bgcolor="#EDF6FF">��˾��ַ��</td>
                        <td width="34%" height="24" bgcolor="#EDF6FF"> 
                          <input class="smallInput" type="text" size="21" name="web" value=<%=rs("web")%>>
                        </td>
                      </tr>
                      <tr> 
                        <td colspan="2" height="24" bgcolor="#EDF6FF"> 
                          <input type="button" value="�� ��" onClick="check()" name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                          <input type="button" class="smallInput"value="�� ��" onClick=javascript:history.go(-1) name="button">
                        </td>
                      </tr>
                    </form>
                  </table>
                        </td>
            </tr>
          </table>
                        <table width="534"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="right"><a href="#top"><img src="images/icon_top.gif" alt="�ص�ҳ�涥��" border="0" align="absmiddle" width="15" height="18"></a> 
                  <a href="#top">�ټ��һ��</a></td>
              </tr>
            </table></td>
        </tr>
      </table>
</td>
  </tr>
</table>

<!--#include file="bottom.htm"-->
</BODY>
</HTML>