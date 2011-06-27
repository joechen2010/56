<!--#include file = "../Inc/Syscode.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from News_NClass order by NClassID asc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("NClassName"))%>","<%= trim(rs("ClassID"))%>","<%= trim(rs("NClassID"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.NClassID.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.NClassID.options[document.myform.NClassID.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
</script>
<title>Untitled Document</title>
	<Script Language=JavaScript>
	// 表单提交客户端检测
	function doSubmit(){
		if (document.myform.Title.value==""){
			alert("新闻标题不能为空！");
			return false;
		}
		// getHTML()为eWebEditor自带的接口函数，功能为取编辑区的内容
		if (eWebEditor1.getHTML()==""){
			alert("新闻内容不能为空！");
			return false;
		}
		document.myform.submit();
	}
	</Script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Inc/Admin_style.css" rel="stylesheet" type="text/css">
</head>

<body>

  <form action="News_AddOK.asp" method="post" name="myform">
	
  <table width="650" align=center cellspacing=3>
    <tr> 
      <td width="71" align="right">标题新闻：</td>
      <td width="564"> 
        <input type="text" name="Title" value="" size="35"></td>
    </tr>
    <tr> 
      <td align="right">新闻栏目：</td>
      <td><%
        sql = "select * from News_Class"
        rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	response.write "请先添加栏目。"
	response.end
	else
%>
        <select name="ClassID" onChange="changelocation(document.myform.ClassID.options[document.myform.ClassID.selectedIndex].value)" size="1">
          <option selected value="<%=trim(rs("ClassID"))%>"><%=trim(rs("ClassName"))%></option>
          <%      dim selclass
         selclass=rs("ClassID")
        rs.movenext
        do while not rs.eof
%>
          <option value="<%=trim(rs("ClassID"))%>"><%=trim(rs("ClassName"))%></option>
          <%
        rs.movenext
        loop
	end if
        rs.close
%>
        </select>
        <select name="NClassID">
          <%sql="select * from News_NClass where ClassID="&selclass
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
%>
          <option selected value="<%=rs("NClassID")%>"><%=rs("NClassName")%></option>
          <% rs.movenext
do while not rs.eof%>
          <option value="<%=rs("NClassID")%>"><%=rs("NClassName")%></option>
          <% rs.movenext
loop
end if
        rs.close
        set rs = nothing
        conn.Close
        set conn = nothing
%>
        </select>
        <font color="#ff6600">**</font></td>
    </tr>
    <tr> 
      <td align="right">新闻内容：</td>
      <td> <input name="Content" type="hidden">
        <iframe ID="eWebEditor1" src="../Editor/ewebeditor.asp?id=Content&style=" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe> 
      </td>
    </tr>
    <tr> 
      <td align="right">标题图片：</td>
      <td><input type="text" name="Picture" id="Picture"> </td>
    </tr>
    <tr> 
      <td align="right">新闻摘自：</td>
      <td><input name="CopyFrom" type="text" id="CopyFrom"></td>
    </tr>
    <tr> 
      <td align="right">新闻属性：</td>
      <td>图文
<input type="checkbox" name="IncludePic" value="yes">
        推荐 
        <input type="checkbox" name="Recommend" value="yes"> </td>
    </tr>
  </table>
	<p align=center><input type=button name=btnSubmit value=" 提 交 " onclick="doSubmit()"> <input type=reset name=btnReset value=" 重 填 "></p>
	</form>


</body>
</html>
