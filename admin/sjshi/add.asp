<!--#include file="../inc/syscode.asp" -->

<head>
<title>添加产品信息</title>
<LINK href="images/Admin.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" topmargin="10">
<table width="90%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#6699ff">
  <form method="POST" action="addproducts.asp">
  <tr align="center"> 
    <td height="27" colspan="2" background="images/Left01.gif"><b>添 加 学 生 信 息</b></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">姓名：</td>
    <td height="30" width="81%">
      <input type="text" name="title" size="20" style="border: 1px outset #C0C0C0">    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">介绍：</td>
      <td height="30" width="81%">
<textarea name="memo" cols="37" rows="7" wrap="physical" style="border: 1px outset #C0C0C0"></textarea>    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td width="19%" height="14" align="right">设计风格：</td>
      <td height="14" width="81%">
<textarea rows="7" name="e_memo" cols="37" style="border: 1px outset #C0C0C0"></textarea>    </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width="19%" height="15" align="right">设计理念:</td>
    <td height="15"><textarea name="linian" cols="37" rows="7" id="linian" style="border: 1px outset #C0C0C0"></textarea></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">图片：</td>
      <td height="30" width="81%">
<iframe name="smallimg" frameborder=0 width=550 height=20 scrolling=no src=smallupload.asp></iframe>
              <br>
              图片路径： 
              <input name="smallimg" type="TEXT" id="smallimg"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100 readonly>    </td>
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
