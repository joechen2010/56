<!--#include file="../inc/syscode.asp" -->
<%session("id")=request("id")
 Sql="select * from shejishi order by id desc"
	'NewsID=Rs(0) NClassID=Rs(1) NClassName=Rs(2) Title=Rs(3) Editor=Rs(4) Hits=Rs(5)  IncludePic=Rs(6)
	'Recommend=Rs(7)
    Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1%>
<head>
<title>��Ӳ�Ʒ��Ϣ</title>
<LINK href="images/Admin.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" topmargin="10">
<table width="90%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#6699ff">
  <form method="POST" action="proadd.asp">
  <tr align="center"> 
    <td height="27" colspan="2" background="images/Left01.gif"><b>�� �� ѧ �� �� Ʒ �� Ϣ</b></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">������</td>
    <td height="30" width="81%"><select name="id" id="id">
	<%while not rs.eof%>
      <option value="<%=rs("id")%>"><%=rs("sjname")%></option>
	  <%rs.movenext
	  wend%>
    </select>    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td width="19%" height="14" align="right">��ƷСͼƬ��</td>
      <td height="14" width="81%">
<iframe name="smallimg" frameborder=0 width=550 height=20 scrolling=no src=smallupload.asp></iframe>
              <br>
              ͼƬ·���� 
              <input name="smallimg" type="TEXT" id="smallimg"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100 readonly>    </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width="19%" height="15" align="right">��Ʒ��ͼƬ:</td>
    <td height="15"><iframe name="BigImg" frameborder=0 width=550 height=20 scrolling=no src=BigUpload.asp></iframe>
      <br>
ͼƬ·����
<input name="BigImg" type="TEXT" id="BigImg"  style="background-color:ffffff;color:000000;border: 1 double" size=34 maxlength=100 readonly></td>
  </tr>
    <tr bgcolor="#FFFFFF" align="center"> 
      <td height="30" colspan="2"> 
        <input type="submit" value=" ȷ ��"  style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="reset" value=" �� �� " name="B2" style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="hidden" name="shibie" value="add">      </td>
  </tr>
  </form>
</table>
</body>
</html>
