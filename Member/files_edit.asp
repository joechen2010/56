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
set rs=conn.execute("select * from file_info where infoID="&infoID)
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
    document.Form1.TypeID.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.Form1.TypeID.options[document.Form1.TypeID.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
</script>


<script LANGUAGE="JavaScript">
function check()
{
if (document.Form1.carnum.value =="") 
{ 
alert("���ƺŲ���Ϊ�գ�"); 
document.Form1.carnum.focus(); 
return (false); 
} 
if (document.Form1.carpp.value =="") 
{ 
alert("Ʒ�Ʋ���Ϊ�գ�"); 
document.Form1.carpp.focus(); 
return (false); 
}

if (document.Form1.cartype.value =="") 
{ 
alert("���ز���Ϊ��"); 
document.Form1.cartype.focus(); 
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
            <a href="http://www.cx56w.com">����������</a> &gt; �Ǽǳ�������</td>
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
                        <div align="center"><font color="#FFFFFF">�޸ĵǼǳ�������</font></div>
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
  <form method="POST" action="files_Edit_save.asp" name="Form1" onSubmit="javascript:return check(this);">
		                  
                    <table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="121" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                        <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">�޸ĳ�������</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">���ƺ�:</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600"> 
						<input name=carnum id=carload value="<%=rs("carnum")%>" maxlength="10">						
                          </font>*</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">Ʒ��: </td>
                        <td width="81%" valign="top" height="22"><font color="#FF6600">							
                          <input name=carpp id=carload value="<%=rs("carpp")%>" maxlength="10">
                          </font>*</td>
                      </tr>					  
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
                          </font>*</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">����:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                          <input name=carLong id=carlong value="<%=rs("carlong")%>" maxlength="4">
                          </font>*</td>
                      </tr>
                      <tr bgcolor="#EDF6FF">
                        <td width="19%" height="22">��������:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
						<input name=time id=carlong value="<%=rs("time")%>" maxlength="10">
                          </font></td>
                      </tr>					  
                      <tr bgcolor="#EDF6FF">
                        <td width="19%" height="22">��Ҫ˾��:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
						<input id=carlong name=siji value="<%=session("siji")%>">
                          </font></td>
                      </tr>					  
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">�ó�״����</td>
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
