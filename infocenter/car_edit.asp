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
    response.write "参数错误，请<a href=""javascript:history.back(-1)"">返回</a>"
	response.end
end if
set rs=conn.execute("select * from info2 where infoID="&infoID)
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
if (document.Form2.province.value =="省份"||document.Form2.province2.value=="省份") 
{ 
alert("请选择区域！"); 
document.Form2.province.focus(); 
return (false); 
} 
 
if (document.Form2.carload.value =="") 
{ 
alert("载重不能为空"); 
document.Form2.carload.focus(); 
return (false); 
}
if (document.Form2.carlong.value =="") 
{ 
alert("车长不能为空"); 
document.Form2.carlong.focus(); 
return (false); 
}
if (document.Form2.content.value.length >50) 
{ 
alert('不能超过50个字!');
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
          <td width="83%" height="38" align="left" id="pos"><a href="http://www.cx56w.com">诚信物流网</a> &gt; 修改车源信息</td>
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
                        <div id="TipAll" > <font color="#FF0000">注：请不要发布重复信息，谢谢合作&nbsp;&nbsp;&nbsp;&nbsp; 
                          </font> </div>                      </td>
                            </tr>
                           </table>
  <form method="POST" action="car_Edit_save.asp" name="Form2" onSubmit="javascript:return check(this);">
		                  
                    <table border="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" cellpadding="5" bordercolordark="#FFFFFF" height="121" cellspacing="1" bgcolor="#B1D4F2">
                      <tr> 
                        <td colspan="2" height="22" background="images/title_bg.gif"><font color="#FFFFFF">发布车源信息</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">出发地:</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600"> 
                          <select id="s1" name="province">
							<option value="省份">省份</option>
                          </select>
                          <select id="s2" name="city">
                            <option>地级市</option>
                          </select>
                          <select id="s3" name="area">
                            <option>市、县级市、县</option>
                          </select>							
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">到达地: </td>
                        <td width="81%" valign="top" height="22"><font color="#FF6600">
                          <select id="t1" name="province2">
                            <option value="省份">省份</option>
                          </select>
                          <select id="t2" name="city2">
                            <option>地级市</option>
                          </select>
                          <select id="t3" name="area2">
                            <option>市、县级市、县</option>
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
                        <td width="19%" height="22">车型:</td>
                        <td width="81%" valign="top" height="22"> 
                        <select name="cartype" id="cartype">
						    <option value="<%=rs("cartype")%>"><%=rs("cartype")%></option>
		                   <option value="有车">有车</option>
		                   <option value="普通车">普通车</option>
		                   <option value="软蓬车">软蓬车</option>
		                   <option value="半挂车">半挂车</option>
		                   <option value="厢式车">厢式车</option>
		                   <option value="冷藏车">冷藏车</option>
		                   <option value="集装箱">集装箱</option>
		                   <option value="平板车">平板车</option>
		                   <option value="起重车">起重车</option>
		                   <option value="自卸车">自卸车</option>
		                   <option value="油罐车">油罐车</option>
		                   <option value="后八轮">后八轮</option>
		                   <option value="前四后八">前四后八</option>
		                   <option value="双桥车">双桥车</option>
		                   <option value="加长挂">加长挂</option>
		                   <option value="半封车">半封车</option>
		                   <option value="高栏车">高栏车</option>
		                   <option value="保温车">保温车</option>
	                </select>						</td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">载重:</td>
                        <td width="81%" valign="top" height="22"> <font color="#FF6600">
                          <input name=carload id=carload value="<%=rs("carload")%>" maxlength="4">
                          吨</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">车长:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                          <input name=carLong id=carlong value="<%=rs("carlong")%>" maxlength="4">
                          米</font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF">
                        <td width="19%" height="22">有效期:</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF"><font color="#FF6600">
                        <select name="validate" id="validate">
						    <option value="<%=rs("validate")%>"><%=rs("validate")%></option>
		                   <option value="5">5天</option>
		                   <option value="15">15天</option>
		                   <option value="30">30天</option>
						   <option value="60">60天</option>
						   <option value="90">90天</option>
						   <option value="120">120天</option>
						   </select>
                          </font></td>
                      </tr>
                      <tr bgcolor="#EDF6FF"> 
                        <td width="19%" height="22">说明：</td>
                        <td width="81%" valign="top" height="22" bgcolor="#EDF6FF">
						<textarea name="content" cols="40" rows="5"><%=rs("content")%></textarea>						</td>
                      </tr>
                      
                      <tr bgcolor="#EDF6FF"> 
                        <td colspan="2" height="19" align="center"> 
						  <input type="hidden" value="<%=rs("infoid")%>" name="infoid">
                          <input type="submit" value="发 布"  name="button">   
						  &nbsp;&nbsp;<input type="button" onClick="javascript:history.back(1)" value="返回">
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
