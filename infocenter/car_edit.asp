<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../Member/check.asp"-->
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
dim rs_c
dim sql_c
dim f
set rs_c=server.createobject("adodb.recordset")
sql_c = "select * from Class_2 order by typeid asc"
rs_c.open sql_c,conn,1,1
%>
<script language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
        f = 0
        do while not rs_c.eof 
        %>
subcat[<%=f%>] = new Array("<%= trim(rs_c("typename"))%>","<%= trim(rs_c("sortid"))%>","<%= trim(rs_c("typeid"))%>");
        <%
        f = f + 1
        rs_c.movenext
        loop
        rs_c.close
        %>
onecount=<%=f%>;

function changelocation(locationid)
    {
    document.Form2.TypeID.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.Form2.TypeID.options[document.Form2.TypeID.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
</script>

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
<link href="../member/images/main.css" rel="stylesheet" type="text/css">
</head>
<BODY>
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="566" align="center" valign="top" bgcolor="#FFFFFF" id="main">
<table width="94%"  border="0" cellspacing="0" cellpadding="3">
      <tr>
          <td width="83%" height="38" align="left" id="pos"><a href="http://www.cx56w.com">����������</a> &gt; �޸ĳ�Դ��Ϣ</td>
          <td width="17%">&nbsp;</td>
      </tr>
    </table>
	  <table width="100%" height="497"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" valign="top"><table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table>
            <table width="534"  border="0" cellpadding="5" cellspacing="1" bgcolor="1E3E93">
              <tr>
              <td valign="top" bgcolor="#FFFFFF">
			  <table width="100%"  border="0">
                             <tr>
                              <td align="left">
                        <div id="TipAll" > <font color="#FF0000">ע���벻Ҫ�����ظ���Ϣ��лл����&nbsp;&nbsp;&nbsp;&nbsp; 
                          </font> </div>                      </td>
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
	                </select>						</td>
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
                        <select name="validate" id="validate">
						    <option value="<%=rs("validate")%>"><%=rs("validate")%></option>
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
						<textarea name="content" cols="40" rows="5"><%=rs("content")%></textarea>						</td>
                      </tr>
                      
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="19" align="center"> 
						  <input type="hidden" value="<%=rs("infoid")%>" name="infoid">
                          <input type="submit" value="�� ��"  name="button">   
						  &nbsp;&nbsp;<input type="button" onClick="javascript:history.back(1)" value="����">
						 </td>
                      </tr>
                    </table>
                          </form>                </td>
            </tr>
          </table>
                        </td>
        </tr>
      </table></td>
  </tr>
</table>

<%
end if
rs.close
set rs=nothing
%>
</BODY>
</HTML>
