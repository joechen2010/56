<!--#include file="../inc/syscode.asp" -->

<head>
<title>��Ӳ�Ʒ��Ϣ</title>
<LINK href="images/Admin.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" topmargin="10">
<table width="90%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#6699ff">
  <form method="POST" action="addproducts.asp">
  <tr align="center"> 
    <td height="27" colspan="2" background="images/Left01.gif"><b>�� �� �� ҵ �� Ϣ</b></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="30" align="right">���ѡ��</td>
    <td height="30">
                          <select name="type">
                            <option>��ҵ����</option>
							<option>�������</option>
                          </select>	</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">��վ���ƣ�</td>
    <td height="30" width="81%">
      <input name="webname" type="text" id="webname" style="border: 1px outset #C0C0C0" size="20">    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="28" width="19%" align="right">��վ��ַ��</td>
      <td width="81%" height="28"><input name="weburl" type="text" id="weburl" style="border: 1px outset #C0C0C0" value="http://" size="20"></td>
  </tr>
    <tr bgcolor="#FFFFFF" align="center"> 
      <td height="30" colspan="2"> 
        <input type="submit" value=" �� �� "  style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="reset" value=" �� �� " name="B2" style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="hidden" name="shibie" value="add">      </td>
  </tr>
  </form>
</table>
</body>
</html>
