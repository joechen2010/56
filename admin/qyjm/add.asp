<!--#include file="../inc/syscode.asp" -->

<head>
<title>添加产品信息</title>
<LINK href="images/Admin.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" topmargin="10">
<table width="90%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#6699ff">
  <form method="POST" action="addproducts.asp">
  <tr align="center"> 
    <td height="27" colspan="2" background="images/Left01.gif"><b>添 加 企 业 信 息</b></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="30" align="right">类别选择：</td>
    <td height="30">
                          <select name="type">
                            <option>企业联盟</option>
							<option>合作伙伴</option>
                          </select>	</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">网站名称：</td>
    <td height="30" width="81%">
      <input name="webname" type="text" id="webname" style="border: 1px outset #C0C0C0" size="20">    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="28" width="19%" align="right">网站网址：</td>
      <td width="81%" height="28"><input name="weburl" type="text" id="weburl" style="border: 1px outset #C0C0C0" value="http://" size="20"></td>
  </tr>
    <tr bgcolor="#FFFFFF" align="center"> 
      <td height="30" colspan="2"> 
        <input type="submit" value=" 添 加 "  style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="reset" value=" 重 置 " name="B2" style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="hidden" name="shibie" value="add">      </td>
  </tr>
  </form>
</table>
</body>
</html>
