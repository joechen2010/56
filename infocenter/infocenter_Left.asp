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
<td height=42 valign=bottom align="center"><a href="infocenter_main.asp?xs=" target="main">显示所有信息</a></td>
</tr>
<tr>
<td height=42 valign=bottom align="center"><a href="infocenter_main.asp?xs=goods" target="main">仅显货源信息</a></td>
</tr>
<tr>
<td height=42 valign=bottom align="center"><a href="infocenter_main.asp?xs=car" target="main">仅显车源信息</a></td>
</tr>
<tr>
<td height=42 valign=bottom align="center">
<TABLE>
    <TR>
      <TD>什么是即时信息？
        <p>　　当有其他会员发布配货信息时，右边的<SPAN id="Label1">货源信息或空车信息将及时显示出来，并显示在最上端．</SPAN></p>
        <p>　　右边的列表每隔 60秒刷新一次,若您想立即看到最新信息,可点击屏幕上端的“既时信息”即可立即刷新页面．</p>
        <p>　　右边的列表是根据您所在的地区筛选出的当地的配货信息，若您还想查看其它地区的信息，可点击屏幕上端的“查找信息”。</p>
        <p>　　可点列表右端的“详情”可查看配货信息的详细内容。   
        <p>　　点击小箭头可折叠或展开左边部分． </p></TD>
    </TR>
</TABLE>
  <p>&nbsp;</p>
  </td>
</tr>
</table>
</body>
</html>