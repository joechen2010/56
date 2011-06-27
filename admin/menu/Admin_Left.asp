<%@language=vbscript codepage=936 %>
<!--#include file="../../inc/config.asp"-->
<html>
<head>
<title>管理页面</title>
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
<table width=158 border="0" align=center cellpadding=0 cellspacing=0> <tr> <td height=42 valign=bottom> 
<img src="images/title.gif" width=158 height=38> </td></tr> </table><table cellpadding=0 cellspacing=0 width=158 align=center> 
<tr> 
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/title_bg_quit.gif id=menuTitle0> 
      &nbsp; &nbsp; <a href="Admin_main.asp" target="main"><b>管理首页</b></a> | <a href=../login/logout.asp target=_top><b>退出</b></a>    </td>
  </tr> <tr> <td style="display:" id='submenu0'> <div class=sec_menu style="width:158"> 
<table cellpadding=0 cellspacing=0 align=center width=130> <tr><td height=20>用户名：<%=session("Adminname")%></td></tr> 
<tr><td height=20>权&nbsp;&nbsp;限：<%
		  select case session("purview")
		  	case 1
				response.write "超级用户"
			case 2
				response.write "高级管理员"
			case 3
				response.write "普通管理员"
		  end select
		  %></td></tr> </table></div><div  style="width:158"> <table cellpadding=0 cellspacing=0 align=center width=130> 
<tr><td height=20></td></tr> </table></div></td></tr> </table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr> 
    <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background="images/Admin_left_1.gif" id=menuTitle1 onClick="showsubmenu(1)" style="cursor:hand;"> 
      <span>新闻管理中心</span> </td>
  </tr>
  <tr> 
    <td style="display:" id='submenu1'> 
      <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr> 
            <td height=20><a href="../News/News_Add.asp" target="main">新增新闻</a> 
              | <a href="../News/News_Manage.asp" target="main">新闻管理</a></td>
          <tr> 
            <td height="20"><a href="../News/NewsClass.asp?action=add" target="main">新闻分栏</a> 
              | <a href="../News/NewsClass_Manage.asp" target="main">分栏管理</a></td>
          </tr>
        </table>
      </div>
      <div  style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20></td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background="images/Admin_left_1.gif" id=menuTitle1 onClick="showsubmenu(10)" style="cursor:hand;"><span>行业分类管理</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu10'><div class=sec_menu style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=130>
        <tr>
          <td height="20"><a href="../sort/sort.asp" target="main">行业分类管理</a></td>
		  </tr>
              </table>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background="images/Admin_left_1.gif" id=menuTitle1 onClick="showsubmenu(17)" style="cursor:hand;"><span>公告管理中心</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu17'><div class=sec_menu style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=130>
        <tr>
          <td height="20"><a href="../News/gong_Add.asp" target="main">添加公告</a>&nbsp;|&nbsp;<a href="../News/gong_Manage.asp" target="main">管理</a></td>
		  </tr>
              </table>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr> 
    <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background="images/Admin_left_2.gif" id=menuTitle1 onClick="showsubmenu(2)" style="cursor:hand;"> 
      <span>注册会员管理中心</span> </td>
  </tr>
  <tr> 
    <td style="display:none" id='submenu2'> 
      <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr> 
            <td height=20><a href="../user/default.asp" target="main">待审会员管理</a></td>
		  </tr>
          <tr> 
            <td height=20><a href="../user/check.asp?jb=0" target="main">免费会员管理</a></td>
		  </tr>
          <tr> 
            <td height=20><a href="../user/check.asp?jb=1" target="main">铜牌会员管理</a></td>
		  </tr>
          <tr> 
            <td height=20><a href="../user/check.asp?jb=2" target="main">银牌会员管理</a></td>
		  </tr>		  		  
          <tr> 
            <td height=20><a href="../user/check.asp?jb=3" target="main">金牌会员管理</a></td>
		  </tr>		  		  
        </table>
      </div>
      <div  style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20></td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background="images/Admin_left_1.gif" id=menuTitle1 onClick="showsubmenu(3)" style="cursor:hand;"><span>发布信息管理中心</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu3'><div class=sec_menu style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=130>
        <tr>
          <td height="20"><a href="../info_center/goods.asp" target="main">货源信息管理</a></td>
	    </tr>
		  <tr>
          <td height="20"><a href="../info_center/cars.asp" target="main">车源信息管理</a></td>
		  </tr>
		  <tr>
          <td height="20"><a href="../info_center/hyzx.asp" target="main">货运专线管理</a></td>
		  </tr>
		  <tr>
          <td height="20"><a href="../info_center/kyzx.asp" target="main">客运专线管理</a></td>
		  </tr>
		  <tr>
          <td height="20"><a href="../info_center/file.asp" target="main">车辆档案管理</a></td>
		  </tr>	
		  <tr>
          <td height="20"><a href="../info_center/repairs.asp" target="main">修理配件管理</a></td>
		  </tr>	
		  <tr>
          <td height="20"><a href="../info_center/kfxx.asp" target="main">库房信息管理</a></td>
		  </tr>	 
		  <tr>
          <td height="20"><a href="../info_center/job.asp" target="main">招聘求职管理</a></td>
		  </tr>	
		  <tr>
          <td height="20"><a href="../info_center/gq.asp" target="main">供求信息管理</a></td>
		  </tr>	
		  <tr>
          <td height="20"><a href="../info_center/banyun.asp" target="main">搬运吊装管理</a></td>
		  </tr>			  		  	         
	  </table>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_4.gif" id=menuTitle1 onClick="showsubmenu(5)" style="cursor:hand;"><span>服务中心</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu5'><div class=sec_menu style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=130>
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=about" target="main">关于我们</A></td>
		</tr>
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=service" target="main">服务条款</A></td>
		</tr>		
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=guide" target="main">网站指南</A></td>
		</tr>
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=charges" target="main">资费标准</A></td>
		</tr>
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=ads" target="main">网站广告</A></td>
		</tr>
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=contact" target="main">联系我们</A></td>
		</tr>
      </table>
    </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr> 
    <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background="images/Admin_left_2.gif" id=menuTitle1 onClick="showsubmenu(37)" style="cursor:hand;"> 
      <span>网站广告管理中心</span> </td>
  </tr>
  <tr> 
    <td style="display:none" id='submenu37'> 
      <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr> 
            <td height=20><a href="../advertise/add.asp" target="main">增加广告</a></td>
          </tr>
          <tr> 
            <td height="20"><a href="../advertise/ad_manage.asp" target="main">管理广告</a></td>
          </tr>
        </table>
      </div>
      <div  style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr>
            <td height=20></td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center> 
<tr> <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_4.gif" id=menuTitle4 onClick="showsubmenu(4)" style="cursor:hand;"> 
<span>系统管理中心</span> </td></tr> <tr> <td style="display:none" id='submenu4'> <div class=sec_menu style="width:158"> 
<table cellpadding=0 cellspacing=0 align=center width=130> <tr> <td HEIGHT="20"><A HREF="../system/User_Manage.asp" TARGET="main">管理员管理</A> 
| <a href="../system/Admin_User.asp?action=add" target="main">添加</a></td></tr><tr><td HEIGHT="20"><A TARGET="main" HREF="../system/system.asp">网站基本设置</A></td></tr></table>
</div><div  style="width:158"> 
<table cellpadding=0 cellspacing=0 align=center width=130> <tr><td height=20></td></tr> 
</table></div></td></tr> </table>
<table cellpadding=0 cellspacing=0 width=158 align=center> 
<tr> <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_9.gif" id=menuTitle9> 
<span>后台管理系统信息</span> </td></tr> <tr> <td> <div class=sec_menu style="width:158"> 
<table cellpadding=0 cellspacing=0 align=center width=130> <tr><td height=20>版权所有：<a href="http://www.yiso.cn/" target=_blank>易商网络</a></td>
</tr>
  <tr>
    <td height=20>设计制作：<a href="http://www.yiso.cn/" target=_blank>易商网络</a></td>
  </tr>
  <tr>
    <td height=20>技术支持：<a href="http://www.yiso.cn/" target=_blank>易商网络</a></td>
  </tr>
</table>
</div></td></tr> </table>
</body>
</html>