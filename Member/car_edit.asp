<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<%
dim infoID
infoID=Request.QueryString("infoID")
if isnumeric(infoID)=0 or infoID="" then
    response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	response.end
end if
set rs=conn.execute("select * from info2 where infoID="&infoID)
if rs.eof and rs.bof then
    response.write "����������<a href=""javascript:history.back(-1)"">����</a>"
	response.end
else
%>

<script language=JavaScript src="../inc/p_c_a.js"></script>
<script language=JavaScript src="../inc/p_c_a2.js"></script>
<script LANGUAGE="JavaScript">
function check()
{
if (document.Form2.province.value =="ʡ��"||document.Form2.province2.value=="ʡ��") 
{ 
alert("��ѡ������"); 
document.Form2.province.focus(); 
return (false); 
} 
 
if (document.Form2.carload.value =="") 
{ 
alert("���ز���Ϊ��"); 
document.Form2.carload.focus(); 
return (false); 
}
if (document.Form2.carlong.value =="") 
{ 
alert("��������Ϊ��"); 
document.Form2.carlong.focus(); 
return (false); 
}
if (document.Form2.content.value.length >50) 
{ 
alert('���ܳ���50����!');
document.Form2.content.focus(); 
return (false); 
}

}
</SCRIPT>
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
.style1 {color: #000000}
.style2 {color: #FFFFFF}
-->
</style>
<link href="images/main.css" rel="stylesheet" type="text/css">
</head>
<BODY>
<!--#include file="head.htm"-->
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">����������</a> &gt; �޸ĳ�Դ��Ϣ</td>
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
                        <div align="center"><font color="#FFFFFF">�޸ĳ�Դ��Ϣ</font></div>
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
                              <td align="left">
                        <div id="TipAll" > <font color="#FF0000">ע���벻Ҫ�����ظ���Ϣ��лл����&nbsp;&nbsp;&nbsp;&nbsp; 
                          </font> </div>
                      </td>
                            </tr>
                           </table>
  <form method="POST" action="car_Edit_save.asp" name="Form2" onSubmit="javascript:return check(this);">
		                  
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
                      <tr bgcolor="#EDF6FF"> 
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
            setup();
          </SCRIPT>
          <SCRIPT language=javascript>
            setup2();
         </SCRIPT>					  
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">����:</td>
                        <td width="81%" valign="top" height="22"> 
                        <select name="cartype" id="cartype">
						    <option value="<%=rs("cartype")%>"><%=rs("cartype")%></option>
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
	                </select>
						</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">����:</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600">
                          <input name=carload id=carload value="<%=rs("carload")%>" maxlength="4">
                          ��</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">����:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                          <input name=carLong id=carlong value="<%=rs("carlong")%>" maxlength="4">
                          ��</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF">
                        <td width="19%" height="22">��Ч��:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                          <input type="text" name="validate" maxlength="3" value="<%=rs("validate")%>">
                        ��</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">˵����</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF">
						<textarea name="content" cols="40" rows="5"><%=rs("content")%></textarea>
						</td>
                      </tr>
                      
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="19" align="center"> 
						  <input type="hidden" value="<%=rs("infoid")%>" name="infoid">
                          <input type="submit" value="�� ��"  name="button">
                       </td>
                      </tr>
                    </table>
                          </form>                </td>
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
<%
end if
rs.close
set rs=nothing
%>
</BODY>
</HTML>
