<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<TITLE>诚信物流网 会员控制中心 - 物流商人的网上家园</TITLE>
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
alert("信息标题不能为空")
document.Form1.title.focus()
document.Form1.title.select()
return
}
if (document.Form1.infotype.value =="") 
{ 
alert("请选择交易类别！"); 
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
            <a href="http://www.cx56w.com">诚信物流网</a> &gt; 发布供求信息</td>
          <td width="17%"><img src="images/icon_onlineservice.gif" width="132" height="32"></td>
      </tr>
    </table>
	  <table width="534"  border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td align="left"><img src="images/icon02.gif" align="absmiddle" width="27" height="19"> 
            <strong> <font color="#cc0000"> <%=session("user")%></font> </strong>，欢迎您回来！</td>
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
                        <div align="center"><font color="#FFFFFF">发布供求信息</font></div>
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
                        <div id="TipAll" > <font color="#FF0000">注：请不要发布重复信息，谢谢合作&nbsp;&nbsp;&nbsp;&nbsp; 
                          </font> </div>
                      </td>
                            </tr>
                           </table>
  <form method="POST" action="info_save.asp" name="Form1">
		                  
                    <table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="121" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                        <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">发布供求信息</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">信息主题：</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600"> 
                          <input id=title size=44 name=title>
                          <br>
                          注意使用合适的标题，勿使用夸张性的词汇或语句。 </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">交易类别： </td>
                        <td width="81%" valign="top" height="22"> 
                          <input type=radio CHECKED value="供应" name=infotype>
                          供应 
                          <input type=radio value="采购" name=infotype>
                          采购(请选择正确的合作方向)<br>
                          <img src="images/warning.gif" border="0" align="absmiddle" width="18" height="17"> 
                          <font color="#FF0000"><strong>特别提醒：</strong><br>
                          错误的合作方向将导致信息不能被审核，直至封禁会员功能或者取消会员资格。</font> </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">关键词：</td>
                        <td width="81%" valign="top" height="22"> 
                          <input id=keyword size=25 name=keyword>
                          <br>
                          <font color="#FF6600">认真选择正确的行业分类，错误的选择将导致信息不能成功发布。 
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">有 效 期：</td>
                        <td width="81%" valign="top" height="22"> 
                          <input type=radio CHECKED value="5" name=validate>
                          5天 
                          <input type=radio value="15" name=validate>
                          15天 
                          <input type=radio value="30" name=validate>
                          30天 
                          <input type=radio value="60" name=validate>
                          60天 
                          <input type=radio value="90" name=validate>
                          90天 
                          <input type=radio value="120" name=validate>
                          120天</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">类别：</td>
                        <td width="81%" valign="top" height="22">
                          <%
        sql_c = "select * from Class_1"
        rs_c.open sql_c,conn,1,1
	if rs_c.eof and rs_c.bof then
	response.write "请先添加栏目。"
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
                        <td width="19%" height="22">内容：</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF">
                          <textarea name="content" cols="45" rows="10" id="content"></textarea>
                          <br>
                          <font color="#FF6600">请尽量详细的描述您的信息要求或者产品资料，这样可以方便其它会员快速找到合适的信息。</font><br>
                          <div  id="myexample"> <img src="images/warning.gif" border="0" align="absmiddle"> 
                            <font color="#FF0000"><strong>特别提醒：</strong><br>
                            请勿在详细说明中提供公司信息或联系方式，包含这些信息的资料将不能通过审核，或者会被管理员删除，多次发现，将取消会员资格。</font></div>
                        </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%">产品图片：</td>
                        <td width="81%" valign="top">&nbsp;<iframe name="InfoUpload" frameborder=0 width=430 height=20 scrolling=no src=Upload.asp?stype=info></iframe> 
                          <br>
                          图片路径： 
                          <input name="ProImg" type="TEXT" id="ProImg"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100>
                        </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="19"> 
                          <input type="button" value="发 布" onClick="check()" name="button">
                          &nbsp;&nbsp;&nbsp;&nbsp; 
                          <input type="button" value="重 填" name="button">
                        </td>
                      </tr>
                    </table>
                          </form>                </td>
            </tr>
          </table>
                        
            <table width="98%"  border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td align="right"><a href="#top"><img src="images/icon_top.gif" alt="回到页面顶部" border="0" align="absmiddle" width="15" height="18"></a> 
                  <a href="#top">再检查一遍</a></td>
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
