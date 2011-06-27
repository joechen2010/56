<!--#include file = "../Inc/Syscode.asp"-->
<% if request("action")="del" then
   page=request("page")
  	if not isempty(request("selAnnounce")) then
     		idlist=request("selAnnounce")
     if instr(idlist,",")>0 then
			dim idarr
			idArr=split(idlist)
			dim id
		for i = 0 to ubound(idarr)
	       		id=clng(idarr(i))
	       		call deleteannounce(id)
		next
		response.Redirect "kfxx.asp?page="&page&""
     		else
			call deleteannounce(clng(idlist))
			response.Redirect "kfxx.asp?page="&page&""
     		end if
  	end if 
	end if
sub deleteannounce(id)
    dim rs_del,sql_del
    set rs_del=server.createobject("adodb.recordset")
    sql_del="delete from kf_info where infoid="&cstr(id)
    conn.execute sql_del
End sub	
if request("pass")="0" then
  conn.execute("update kf_info set pass=false where infoid="&request("id"))
  response.Redirect "?page="&request("page")&""
  elseif request("pass")="1" then
   conn.execute("update kf_info set pass=true where infoid="&request("id"))
  response.Redirect "?page="&request("page")&"" 
end if
'delete -----------
'if  Request.QueryString("del")<>"" then
   'delsql="delete from kfxx where id="&Request.QueryString("del")
  ' conn.execute(delsql)
'end if
'delete end---------------
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../inc/Admin_Style.css">
<script language="javascript">
<!--

function mOvr(src,clrOver){
	if (!src.contains(event.fromElement)) {
		//src.style.cursor = 'hand';
		src.bgColor = clrOver;
	}
}
function mOut(src,clrIn)  {
	if (!src.contains(event.toElement)) {
		src.style.cursor = 'default';
		src.bgColor = clrIn;
	}
}
function mClk1(src,id) {
		window.location .href =src;
 }
function ConfirmDel()
{
   if(confirm("确定要删除该项记录吗？        \n点击确定删除！\n点击取消返回！"))
     return true;
   else
     return false;
	 
}


function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled==false)
       e.checked = form.chkAll.checked;
    }
  }
function cform(){
 if(!confirm("Are you sure?"))
 return false;
}  
//-->
</script>
<script language="javascript">
function openwindow(url,width,height) { 	left1 = (screen.width-800)/2; 	top1 = (screen.height-350)/2; 	window.open(url,"","width=" + width + ",height=" + height + ",left=" + left1.toString() + ",top=" + top1.toString()); }

</script>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#e1eefd">
<div align="center"> 
  <table width="96%" border="0" align="center" cellspacing="0" class="px14">
    <tr> 
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#6699ff" height="30"><font color="#FFFFFF"> 库房信息管理</font></td>
    </tr>
  </table>
   <br>
   <form method="post" name="myform" action="kfxx.asp">
   <% 
	  set rs=Server.Createobject("ADODB.Recordset")
      sql="select  * from kf_info order by infoid desc"     
       rs.open sql,conn,1,3
	  if rs.recordcount=0 then
	  Response.Write "无内容"
	  Response.end
	  else
          rs.pagesize=20
          if request("page")="" then
           page=0
           else
           page=clng(request("page"))
          end if
           if page<1 then page=1
           if page>rs.pagecount then page=rs.pagecount
          rs.AbsolutePage=page
		%>
  <table width="96%" border="0" cellspacing="1">
    <tr bgcolor="#CCCCCC"> 
    <td width="246" height="25" bgcolor="#e1eefd" align="center">信息类型</td>
    <td width="331" height="25" bgcolor="#e1eefd" align="center">库房类型</td>
    <td width="213" height="25" bgcolor="#e1eefd" align="center">仓库地点</td>
	<td width="331" height="25" bgcolor="#e1eefd" align="center">面积</td>
	<td width="331" height="25" bgcolor="#e1eefd" align="center">价格</td>
	<td width="331" height="25" bgcolor="#e1eefd" align="center">发布时间</td>
	<td width="213" height="25" bgcolor="#e1eefd" align="center">审核</td>
    <td width="167"  bgcolor="#e1eefd" align="center">删除</td>
    </tr>
	 <% for ipage=1 to rs.pagesize%>
   <tr bgcolor="#F6F6F6"  onMouseOut="mOut(this,'#F6F6F6');" onMouseOver="mOvr(this,'#EBFFBB');"> 
      <td>&nbsp;<a href="javascript:openwindow('../../info/detail.asp?infoid=<%=rs("InfoID")%>',500,400)">	  
	  <%=rs("infotype")%></a></td>
      <td  align="center"><%=rs("kftype")%></td>
      <td align="center"><%=rs("city")%>&nbsp;<%=rs("area")%></td>
      <td align="center"><%=rs("mji")%>平方米</td>
      <td  align="center"><%=rs("prices")%>元/平方米</td>
	  <td align="center"><%=rs("addtime")%></td>
	<td width="213" align="center">
	  <%
	  if rs("pass")=true then
	    response.Write "<a href='?pass=0&page="&page&"&id="&rs("infoid")&"'>已审核</a>"
	   else
	    response.write "<a href='?pass=1&page="&page&"&id="&rs("infoid")&"'>未审核</a>"	
	  end if
	  %>	
	</td>	  
      <td align="center"><input type='checkbox' name='selAnnounce' value='<%=cstr(rs("InfoID"))%>'>
	  <!--<a href="?del=<%'=rs("id")%>"  onClick="return cform()">删除</a>--></td>
    </tr>
	<%rs.movenext
            if rs.eof then exit for
            next%>
  </table> 
  <table width="96%" height="21" border="0" cellpadding="0" cellspacing="1" background="images/bg-xg.gif" bgcolor="#EFEFEF">
    <tr> 
    <td width="37%" align="center" bgcolor="#e7e7e7"> 
	 <input type="hidden" name="action" value="del">
	 <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #eeeeee;"><label for="chkAll">全部选中</label>&nbsp;&nbsp;<input type="submit" name="sc" value="删&nbsp;除&nbsp;所&nbsp;选" onClick="return cform()"></td>
	 </tr>
	 <tr><td width="63%" align="right" bgcolor="#e1eefd">	每页显示<font color="#FF0000"><%=rs.pagesize%></font>条&nbsp;&nbsp;	
	     共有<font color="#FF0000"><%=rs.recordcount%></font>条信息	&nbsp;&nbsp;
			当前<font color="#FF0000"><%=page%></font>/<%=rs.pagecount%>
              <% if page<>1 then %>
              <a href="?page=1">首页</a>&nbsp;<a href=?page=<%=(page-1)%>>上一页</a> 
              <%end if%>
              <% if page<>rs.pagecount then %>
              <a href=?page=<%=(page+1)%>>下一页</a>&nbsp;<a href="?page=<%=rs.pagecount%>">最后一页</a> 
              <%end if%>
			 <input type="text" name="page" value="<%=page%>" size="4">
			 <input type="submit" value="跳&nbsp;转">
        </td>
    </tr>
  </table> 
	  <% 	rs.close
			set rs=nothing
			End If %>	
  </form>
</div>
<br>
</body>
</html>
<% 
conn.close
set conn=nothing %>