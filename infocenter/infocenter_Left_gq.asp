<%@language=vbscript codepage=936 %>
<!--#include file="../inc/config.asp"-->
<html>
<head>
<title>信息中心</title>
<style type=text/css>
body  { background:#799AE1; margin:0px; font:9pt 宋体; }
table  { border:0px; }
td  { font:normal 12px 宋体; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px 宋体; color:#000000; text-decoration:none; }
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
<td align="center" valign=bottom>
  <ul>
    <LI>右边的列表列出的是您自己已发布的全部供求信息．   
    <LI>点击“操作”可修改并重新发布您已经发布的供求信息（发布时间自动改为当日）。   
    <LI>也可删除您已经发布的供求信息。   
    <LI>若要发布信息可点击列表下面的“发布供求信息”。   
    <LI>选择信息类型和行业类目若不想精确到细类，第四个下拉框可不选择。   
    <LI>可以修改产品或服务名称框中的内容，但必须与行业类目相关．   
    <LI>有效期的作用是过了有效期的天数后，其它会员将看不到这条信息，但您自己能看到，并能重新发布。   
    <LI>在“详细说明”栏中可填入具体的内容，不能超过100字。   
    <LI>全部必填栏填好后点击“发布”即可。</LI>
  </ul>
  <H2 align="center">&nbsp;</H2>
  </td>
</tr>
</table>
</body>
</html>