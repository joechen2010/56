<%@language=vbscript codepage=936 %>
<!--#include file="../inc/config.asp"-->
<html>
<head>
<title>��Ϣ����</title>
<style type=text/css>
body  { background:#799AE1; margin:0px; font:9pt ����; }
table  { border:0px; }
td  { font:normal 12px ����; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px ����; color:#000000; text-decoration:none; }
a:hover  { color:#cc0000;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#cc0000; font-weight:bold; }
</style>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
</head>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=158 border="0" align=center cellpadding=0 cellspacing=0>
<tr> 
<td height=42 valign=bottom align="center"><a href="infocenter_main.asp?xs=" target="main">��ʾ������Ϣ</a></td>
</tr>
<tr>
<td height=42 valign=bottom align="center"><a href="infocenter_main.asp?xs=goods" target="main">���Ի�Դ��Ϣ</a></td>
</tr>
<tr>
<td height=42 valign=bottom align="center"><a href="infocenter_main.asp?xs=car" target="main">���Գ�Դ��Ϣ</a></td>
</tr>
<tr>
<td height=42 valign=bottom align="center">
<TABLE>
    <TR>
      <TD>ʲô�Ǽ�ʱ��Ϣ��
        <p>��������������Ա���������Ϣʱ���ұߵ�<SPAN id="Label1">��Դ��Ϣ��ճ���Ϣ����ʱ��ʾ����������ʾ�����϶ˣ�</SPAN></p>
        <p>�����ұߵ��б�ÿ�� 60��ˢ��һ��,��������������������Ϣ,�ɵ����Ļ�϶˵ġ���ʱ��Ϣ����������ˢ��ҳ�森</p>
        <p>�����ұߵ��б��Ǹ��������ڵĵ���ɸѡ���ĵ��ص������Ϣ����������鿴������������Ϣ���ɵ����Ļ�϶˵ġ�������Ϣ����</p>
        <p>�����ɵ��б��Ҷ˵ġ����顱�ɲ鿴�����Ϣ����ϸ���ݡ�   
        <p>�������С��ͷ���۵���չ����߲��֣� </p></TD>
    </TR>
</TABLE>
  <p>&nbsp;</p>
  </td>
</tr>
</table>
</body>
</html>