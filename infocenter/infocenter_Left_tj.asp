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
  <H2 align="center">什么是推荐信息？</H2>
  <p>推荐信息是系统根据您在本站分布的配货信息，在数据库中自动检索出与之相匹配的对应信息，</p>
  <p>匹配条件是出发地和目的地相同或相近（市级匹配）的有效期内的信息。</p>
  <p>例如，您在今天已经发布了一条从长春到北京的货源信息，有限期为５天，那么在今后的５天内，推荐信息列出所有其它会员发布的长春到北京的有效期内的车源信息，供您参考．</p>
  <p>至于车型、运价、发车时间是否满足您的要求，您可以点击推荐信息列表中的“祥情”查看详细内容，并可以根据对方的联系方式与之进一步洽谈。</p>
  <p>同样，如果您发布了车源信息，推荐信息也将列出所有与之相匹配的货源信息。</p></td>
</tr>
</table>
</body>
</html>