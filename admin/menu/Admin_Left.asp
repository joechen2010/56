<%@language=vbscript codepage=936 %>
<!--#include file="../../inc/config.asp"-->
<html>
<head>
<title>����ҳ��</title>
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
<table width=158 border="0" align=center cellpadding=0 cellspacing=0> <tr> <td height=42 valign=bottom> 
<img src="images/title.gif" width=158 height=38> </td></tr> </table><table cellpadding=0 cellspacing=0 width=158 align=center> 
<tr> 
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/title_bg_quit.gif id=menuTitle0> 
      &nbsp; &nbsp; <a href="Admin_main.asp" target="main"><b>������ҳ</b></a> | <a href=../login/logout.asp target=_top><b>�˳�</b></a>    </td>
  </tr> <tr> <td style="display:" id='submenu0'> <div class=sec_menu style="width:158"> 
<table cellpadding=0 cellspacing=0 align=center width=130> <tr><td height=20>�û�����<%=session("Adminname")%></td></tr> 
<tr><td height=20>Ȩ&nbsp;&nbsp;�ޣ�<%
		  select case session("purview")
		  	case 1
				response.write "�����û�"
			case 2
				response.write "�߼�����Ա"
			case 3
				response.write "��ͨ����Ա"
		  end select
		  %></td></tr> </table></div><div  style="width:158"> <table cellpadding=0 cellspacing=0 align=center width=130> 
<tr><td height=20></td></tr> </table></div></td></tr> </table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr> 
    <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background="images/Admin_left_1.gif" id=menuTitle1 onClick="showsubmenu(1)" style="cursor:hand;"> 
      <span>���Ź�������</span> </td>
  </tr>
  <tr> 
    <td style="display:" id='submenu1'> 
      <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr> 
            <td height=20><a href="../News/News_Add.asp" target="main">��������</a> 
              | <a href="../News/News_Manage.asp" target="main">���Ź���</a></td>
          <tr> 
            <td height="20"><a href="../News/NewsClass.asp?action=add" target="main">���ŷ���</a> 
              | <a href="../News/NewsClass_Manage.asp" target="main">��������</a></td>
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
    <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background="images/Admin_left_1.gif" id=menuTitle1 onClick="showsubmenu(10)" style="cursor:hand;"><span>��ҵ�������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu10'><div class=sec_menu style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=130>
        <tr>
          <td height="20"><a href="../sort/sort.asp" target="main">��ҵ�������</a></td>
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
    <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background="images/Admin_left_1.gif" id=menuTitle1 onClick="showsubmenu(17)" style="cursor:hand;"><span>�����������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu17'><div class=sec_menu style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=130>
        <tr>
          <td height="20"><a href="../News/gong_Add.asp" target="main">��ӹ���</a>&nbsp;|&nbsp;<a href="../News/gong_Manage.asp" target="main">����</a></td>
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
      <span>ע���Ա��������</span> </td>
  </tr>
  <tr> 
    <td style="display:none" id='submenu2'> 
      <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr> 
            <td height=20><a href="../user/default.asp" target="main">�����Ա����</a></td>
		  </tr>
          <tr> 
            <td height=20><a href="../user/check.asp?jb=0" target="main">��ѻ�Ա����</a></td>
		  </tr>
          <tr> 
            <td height=20><a href="../user/check.asp?jb=1" target="main">ͭ�ƻ�Ա����</a></td>
		  </tr>
          <tr> 
            <td height=20><a href="../user/check.asp?jb=2" target="main">���ƻ�Ա����</a></td>
		  </tr>		  		  
          <tr> 
            <td height=20><a href="../user/check.asp?jb=3" target="main">���ƻ�Ա����</a></td>
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
    <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background="images/Admin_left_1.gif" id=menuTitle1 onClick="showsubmenu(3)" style="cursor:hand;"><span>������Ϣ��������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu3'><div class=sec_menu style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=130>
        <tr>
          <td height="20"><a href="../info_center/goods.asp" target="main">��Դ��Ϣ����</a></td>
	    </tr>
		  <tr>
          <td height="20"><a href="../info_center/cars.asp" target="main">��Դ��Ϣ����</a></td>
		  </tr>
		  <tr>
          <td height="20"><a href="../info_center/hyzx.asp" target="main">����ר�߹���</a></td>
		  </tr>
		  <tr>
          <td height="20"><a href="../info_center/kyzx.asp" target="main">����ר�߹���</a></td>
		  </tr>
		  <tr>
          <td height="20"><a href="../info_center/file.asp" target="main">������������</a></td>
		  </tr>	
		  <tr>
          <td height="20"><a href="../info_center/repairs.asp" target="main">�����������</a></td>
		  </tr>	
		  <tr>
          <td height="20"><a href="../info_center/kfxx.asp" target="main">�ⷿ��Ϣ����</a></td>
		  </tr>	 
		  <tr>
          <td height="20"><a href="../info_center/job.asp" target="main">��Ƹ��ְ����</a></td>
		  </tr>	
		  <tr>
          <td height="20"><a href="../info_center/gq.asp" target="main">������Ϣ����</a></td>
		  </tr>	
		  <tr>
          <td height="20"><a href="../info_center/banyun.asp" target="main">���˵�װ����</a></td>
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
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_4.gif" id=menuTitle1 onClick="showsubmenu(5)" style="cursor:hand;"><span>��������</span> </td>
  </tr>
  <tr>
    <td style="display:none" id='submenu5'><div class=sec_menu style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=130>
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=about" target="main">��������</A></td>
		</tr>
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=service" target="main">��������</A></td>
		</tr>		
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=guide" target="main">��վָ��</A></td>
		</tr>
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=charges" target="main">�ʷѱ�׼</A></td>
		</tr>
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=ads" target="main">��վ���</A></td>
		</tr>
        <tr>
          <td height=20><a href="../service/service_Add.asp?action=contact" target="main">��ϵ����</A></td>
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
      <span>��վ����������</span> </td>
  </tr>
  <tr> 
    <td style="display:none" id='submenu37'> 
      <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr> 
            <td height=20><a href="../advertise/add.asp" target="main">���ӹ��</a></td>
          </tr>
          <tr> 
            <td height="20"><a href="../advertise/ad_manage.asp" target="main">������</a></td>
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
<span>ϵͳ��������</span> </td></tr> <tr> <td style="display:none" id='submenu4'> <div class=sec_menu style="width:158"> 
<table cellpadding=0 cellspacing=0 align=center width=130> <tr> <td HEIGHT="20"><A HREF="../system/User_Manage.asp" TARGET="main">����Ա����</A> 
| <a href="../system/Admin_User.asp?action=add" target="main">���</a></td></tr><tr><td HEIGHT="20"><A TARGET="main" HREF="../system/system.asp">��վ��������</A></td></tr></table>
</div><div  style="width:158"> 
<table cellpadding=0 cellspacing=0 align=center width=130> <tr><td height=20></td></tr> 
</table></div></td></tr> </table>
<table cellpadding=0 cellspacing=0 width=158 align=center> 
<tr> <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/Admin_left_9.gif" id=menuTitle9> 
<span>��̨����ϵͳ��Ϣ</span> </td></tr> <tr> <td> <div class=sec_menu style="width:158"> 
<table cellpadding=0 cellspacing=0 align=center width=130> <tr><td height=20>��Ȩ���У�<a href="http://www.yiso.cn/" target=_blank>��������</a></td>
</tr>
  <tr>
    <td height=20>���������<a href="http://www.yiso.cn/" target=_blank>��������</a></td>
  </tr>
  <tr>
    <td height=20>����֧�֣�<a href="http://www.yiso.cn/" target=_blank>��������</a></td>
  </tr>
</table>
</div></td></tr> </table>
</body>
</html>