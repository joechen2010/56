<%@ codepage ="936" %>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<%
dim ProID
ProID=Request.QueryString("ProID")
if isnumeric(ProID)=0 or ProID="" then
    response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	response.end
end if
set rs=conn.execute("select * from Spzs where ProID="&ProID)
if rs.eof and rs.bof then
    response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
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
if (document.Form1.BearinNo.value=="")
{
alert("产品型号不能为空")
document.Form1.BearinNo.focus()
document.Form1.BearinNo.select()
return
}
if (document.Form1.Brand.value =="") 
{ 
alert("请选择交易类别！"); 
document.Form1.Brand.focus(); 
return (false); 
} 
if (eWebEditor1.getHTML()=="")
{
alert("信息内容不能为空")
document.Form1.Introduce.focus()
document.Form1.Introduce.select() 
return
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
          <td align="center" valign="top"><table width="534" border="0" align="center" cellpadding="0" cellspacing="0">
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
  <form method="POST" action="Pro_Edit_save.asp" name="Form1">
		                  
                    <table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="121" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                        <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">发布供求信息</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">型号：</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600"> 
                          <input id="BearinNo" name="BearinNo" value="<%=rs("BearinNo")%>">
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">品牌： </td>
                        <td width="81%" valign="top" height="22"><font color="#FF6600">
                          <input id="Brand" name="Brand" value="<%=rs("Brand")%>">
                          </font> </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">数量：</td>
                        <td width="81%" valign="top" height="22">
                          <input id="Num" size="25" name="Num" value="<%=rs("Num")%>">
                        </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">单价：</td>
                        <td width="81%" valign="top" height="22"><font color="#FF6600">
                          <input id="Price" name="Price" value="<%=rs("Price")%>">
                          </font> </td>
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
                            <%
do while not rs_c.eof

     
%>
                            <option <%if trim(rs("sortid"))= trim(rs_c("SortID")) then response.write "selected"%> value="<%=trim(rs_c("SortID"))%>"><%=trim(rs_c("sort"))%></option>
                            <%
        rs_c.movenext
        loop
	end if
        rs_c.close
%>
                          </select>
                          <select name="TypeID">
                            <%sql_c="select * from Class_2 where SortID="&rs("sortid")
rs_c.open sql_c,conn,1,1
if not(rs_c.eof and rs_c.bof) then
do while not rs_c.eof%>
                            <option <%if rs_c("typeid")=rs("typeid") then response.write "selected"%> value="<%=rs_c("TypeID")%>"><%=rs_c("TypeName")%></option>
                            <% rs_c.movenext
loop
end if
        rs_c.close
        set rs_c = nothing
%>
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EDF6FF">
                        <td width="19%" height="22">现货地点：</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                          <input id="Nowplace" name="Nowplace" value="<%=rs("Nowplace")%>">
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">产地：</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                          <input id="MadePlace" name="MadePlace" value="<%=rs("MadePlace")%>">
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">说明：</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF">&nbsp; 
                          <input type="hidden" name="ProID" value="<%=rs("ProID")%>">
                        </td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="22"> 
                          <input name="Introduce" type="hidden" value="<%=Server.HtmlEncode(Rs("Introduce"))%>">
                          <iframe id="eWebEditor1" src="../Editor/ewebeditor.asp?id=Introduce&style=" frameborder="0" scrolling="no" width="550" height="350"></iframe></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%">产品图片：</td>
                        <td width="81%" valign="top">&nbsp;<iframe name="InfoUpload" frameborder=0 width=430 height=20 scrolling=no src=Upload.asp?stype=Pro></iframe> 
                          <br>
                          图片路径： 
                          <input name="ProImg" type="TEXT" id="ProImg"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100 value="<%=rs("ProImg")%>">
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
                        <table width="534"  border="0" cellspacing="0" cellpadding="5">
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
<%
end if
rs.close
set rs=nothing
%>
</BODY>
</HTML>
