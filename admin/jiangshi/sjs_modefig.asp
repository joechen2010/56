<!--#include file="../inc/syscode.asp" -->
<%session("id")=request("id")
 Sql="select * from jsfc where id="&request("id")&""
	'NewsID=Rs(0) NClassID=Rs(1) NClassName=Rs(2) Title=Rs(3) Editor=Rs(4) Hits=Rs(5)  IncludePic=Rs(6)
	'Recommend=Rs(7)
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1%>
<head>
<title>添加产品信息</title>
<LINK href="images/Admin.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" topmargin="10">
<table width="90%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#6699ff">
  <form method="POST" action="modfiy.asp">
  <tr align="center"> 
    <td height="27" colspan="2" background="images/Left01.gif"><b>修 改 学 生 信 息</b></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">姓名：</td>
    <td height="30" width="81%">
      <input name="title" type="text" style="border: 1px outset #C0C0C0" value="<%=rs("jsname")%>" size="20">    </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="28" align="right">出生日期:</td>
    <td height="28"><input name="age" type="text" id="age" style="border: 1px outset #C0C0C0" value="<%=rs("age")%>" size="20"></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="29" align="right">所任职务:</td>
    <td height="29"><input name="zhiwu" type="text" id="zhiwu" style="border: 1px outset #C0C0C0" value="<%=rs("zhiwu")%>" size="20"></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="30" align="right">职称:</td>
    <td height="30"><input name="zhicheng" type="text" id="zhicheng" style="border: 1px outset #C0C0C0" value="<%=rs("zhicheng")%>" size="20"></td>
  </tr>
  
  <tr bgcolor="#FFFFFF"> 
      <td width="19%" height="14" align="right">介绍：</td>
      <td height="14" width="81%">
<textarea rows="7" name="e_memo" cols="37" style="border: 1px outset #C0C0C0"><%=rs("produ")%></textarea>    </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width="19%" height="15" align="right">所获奖项:</td>
    <td height="15"><textarea name="linian" cols="37" rows="7" id="linian" style="border: 1px outset #C0C0C0"><%=rs("jiang")%></textarea></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">图片：</td>
      <td height="30" width="81%">
<iframe name="smallimg" frameborder=0 width=550 height=20 scrolling=no src=smallupload.asp></iframe>
              <br>
              图片路径： 
              <input name="smallimg" type="TEXT" id="smallimg"  style="background-color:ffffff;color:000000;border: 1 double" value="<%=rs("smallpic")%>" size=34 maxlength=100 readonly>    </td>
  </tr>
    <tr bgcolor="#FFFFFF" align="center"> 
      <td height="30" colspan="2"> 
        <input type="submit" value=" 修 改"  style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="reset" value=" 重 置 " name="B2" style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="hidden" name="shibie" value="add">      </td>
  </tr>
  </form>
</table>
</body>
</html>
