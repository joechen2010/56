<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<TITLE>���������� ��Ա�������� - �������˵����ϼ�԰</TITLE>
<%
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
if (document.Form1.title.value=="")
{
alert("��Ϣ���ⲻ��Ϊ��")
document.Form1.title.focus()
document.Form1.title.select()
return
}
if (document.Form1.infotype.value =="") 
{ 
alert("��ѡ�������"); 
document.Form1.infotype.focus(); 
return (false); 
} 

document.Form1.submit()
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
<table width="778"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="160" height="566" valign="top" bgcolor="#2C68B1">
<!--#include file="left.htm"-->
	</td>
    <td id="main" align="center" valign="top" bgcolor="#FFFFFF">
      <table width="534"  border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="83%" height="38" align="left" id="pos"><img src="images/icon03.gif" width="15" height="11"> 
            <a href="http://www.cx56w.com">����������</a> &gt; ����������Ϣ</td>
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
          <td align="center" valign="top">
            <table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="534"  border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="12" height="21" align="left" valign="middle" bgcolor="#FFFFFF" ><img src="images/title_1.jpg" width="12" height="25"></td>
                      <td valign="middle" background="images/title_2.jpg" bgcolor="#FFFFFF" >
                        <div align="center"><font color="#FFFFFF">����������Ϣ</font></div>
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
  <form method="POST" action="info_save.asp" name="Form1">
		                  
                    <table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="121" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                        <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">����������Ϣ</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">��Ϣ���⣺</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600"> 
                          <input id=title size=44 name=title>
                          <br>
                          ע��ʹ�ú��ʵı��⣬��ʹ�ÿ����ԵĴʻ����䡣 </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">������� </td>
                        <td width="81%" valign="top" height="22"> 
                          <input type=radio CHECKED value="��Ӧ" name=infotype>
                          ��Ӧ 
                          <input type=radio value="�ɹ�" name=infotype>
                          �ɹ�(��ѡ����ȷ�ĺ�������)<br>
                          <img src="images/warning.gif" border="0" align="absmiddle" width="18" height="17"> 
                          <font color="#FF0000"><strong>�ر����ѣ�</strong><br>
                          ����ĺ������򽫵�����Ϣ���ܱ���ˣ�ֱ�������Ա���ܻ���ȡ����Ա�ʸ�</font> </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">�ؼ��ʣ�</td>
                        <td width="81%" valign="top" height="22"> 
                          <input id=keyword size=25 name=keyword>
                          <br>
                          <font color="#FF6600">����ѡ����ȷ����ҵ���࣬�����ѡ�񽫵�����Ϣ���ܳɹ������� 
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">�� Ч �ڣ�</td>
                        <td width="81%" valign="top" height="22"> 
                          <input type=radio CHECKED value="5" name=validate>
                          5�� 
                          <input type=radio value="15" name=validate>
                          15�� 
                          <input type=radio value="30" name=validate>
                          30�� 
                          <input type=radio value="60" name=validate>
                          60�� 
                          <input type=radio value="90" name=validate>
                          90�� 
                          <input type=radio value="120" name=validate>
                          120��</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">���</td>
                        <td width="81%" valign="top" height="22">
                          <%
        sql_c = "select * from Class_1"
        rs_c.open sql_c,conn,1,1
	if rs_c.eof and rs_c.bof then
	response.write "���������Ŀ��"
	response.end
	else
%>
                          <select name="SortID" onChange="changelocation(document.Form1.SortID.options[document.Form1.SortID.selectedIndex].value)" size="1">
                            <option selected value="<%=trim(rs_c("SortID"))%>"><%=trim(rs_c("sort"))%></option>
                            <%      dim selclass
         selclass=rs_c("SortID")
        rs_c.movenext
        do while not rs_c.eof
%>
                            <option value="<%=trim(rs_c("SortID"))%>"><%=trim(rs_c("sort"))%></option>
                            <%
        rs_c.movenext
        loop
	end if
        rs_c.close
%>
                          </select>
                          <select name="TypeID">
                            <%sql_c="select * from Class_2 where SortID="&selclass
rs_c.open sql_c,conn,1,1
if not(rs_c.eof and rs_c.bof) then
%>
                            <option selected value="<%=rs_c("TypeID")%>"><%=rs_c("TypeName")%></option>
                            <% rs_c.movenext
do while not rs_c.eof%>
                            <option value="<%=rs_c("TypeID")%>"><%=rs_c("TypeName")%></option>
                            <% rs_c.movenext
loop
end if
        rs_c.close
        set rs_c = nothing
        conn.Close
        set conn = nothing
%>
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">���ݣ�</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF">
                          <textarea name="content" cols="45" rows="10" id="content"></textarea>
                          <br>
                          <font color="#FF6600">�뾡����ϸ������������ϢҪ����߲�Ʒ���ϣ��������Է���������Ա�����ҵ����ʵ���Ϣ��</font><br>
                          <div  id="myexample"> <img src="images/warning.gif" border="0" align="absmiddle"> 
                            <font color="#FF0000"><strong>�ر����ѣ�</strong><br>
                            ��������ϸ˵�����ṩ��˾��Ϣ����ϵ��ʽ��������Щ��Ϣ�����Ͻ�����ͨ����ˣ����߻ᱻ����Աɾ������η��֣���ȡ����Ա�ʸ�</font></div>
                        </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%">��ƷͼƬ��</td>
                        <td width="81%" valign="top">&nbsp;<iframe name="InfoUpload" frameborder=0 width=430 height=20 scrolling=no src=Upload.asp?stype=info></iframe> 
                          <br>
                          ͼƬ·���� 
                          <input name="ProImg" type="TEXT" id="ProImg"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100>
                        </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="19"> 
                          <input type="button" value="�� ��" onClick="check()" name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                          <input type="button" value="�� ��" name="button">
                        </td>
                      </tr>
                    </table>
                          </form>                </td>
            </tr>
          </table>
                        
            <table width="98%"  border="0" cellspacing="0" cellpadding="5">
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
