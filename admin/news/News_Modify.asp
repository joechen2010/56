<!--#include file = "../Inc/Syscode.asp"-->
<%Dim ClassID,NewsID,RsNews,SqlStr
ClassID=Request("ClassID")
NewsID=Request("NewsID")
Dim sNewsID,sNClassID,sNClassName,sTitle,sContent,sPicture,sCopyFrom,sEditor,sIncludePic,sHits,sRecommend,sRegTime
set RsNews = Server.CreateObject("Adodb.Recordset")
    SqlStr = "Select NewsID,NClassID,(Select NClassName from News_NClass where NClassID=A.NClassID) as NClassName,Title,Content,Picture,CopyFrom,Editor,IncludePic,Hits,Recommend,RegTime,ClassID,(select ClassName from News_Class where ClassID=a.ClassID) as ClassName from NewsData A where NewsID =" &NewsID
	RsNews.open SqlStr,Conn,1,1
	
%>
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
<title></title>
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
<link href="../Inc/admin_style.css" rel="stylesheet" type="text/css">
</head>

<body>

  <form action="News_ModifyOK.asp" method="post" name="myform">
  <table width="650" align=center cellspacing=3>
    <tr> 
      <td width="71" align="right">标题新闻：</td>
      <td width="564"> <input type="text" name="Title" value="<%=RsNews(3)%>" size="35"></td>
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
          <option selected value="<%=trim(rsnews(12))%>"><%=trim(rsnews(13))%></option>
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
	<option selected value="<%=rsnews(1)%>"><%=rsnews(2)%></option>
          <%sql="select * from News_NClass where ClassID="&selclass
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
%>
          
          <% rs.movenext
do while not rs.eof%>
          <option value="<%=rs("NClassID")%>"><%=rs("NClassName")%></option>
          <% rs.movenext
loop
end if
        rs.close
        set rs = nothing
%>
        </select>
        <font color="#ff6600">**</font></td>
    </tr>
    <tr> 
      <td align="right">新闻内容：</td>
      <td> <input name="Content" type="hidden" value="<%=Server.HtmlEncode(RsNews(4))%>"> 
        <iframe id="eWebEditor1" src="../Editor/ewebeditor.asp?id=Content&style=" frameborder="0" scrolling="no" width="550" height="350"></iframe> 
      </td>
    </tr>
    <tr> 
      <td align="right">标题图片：</td>
      <td><input type="text" name="Picture" id="Picture" value="<%=RsNews(5)%>"> 
      </td>
    </tr>
    <tr> 
      <td align="right">新闻摘自：</td>
      <td><input name="CopyFrom" type="text" id="CopyFrom" value="<%=RsNews(6)%>"></td>
    </tr>
    <tr> 
      <td align="right">新闻属性：</td>
      <td>图文 
        <input type="checkbox" name="IncludePic" <%if RsNews(1)=true then response.write "checked" end if%> value="yes">
        推荐 
        <input name="Recommend" type="checkbox" <%if RsNews("Recommend")=true then response.write "checked" end if %> value="yes"> 
      </td>
    </tr>
  </table><input name="NewsID" type="hidden" value="<%=RsNews(0)%>">
  <p align=center><input type=button name=btnSubmit value=" 提 交 " onClick="doSubmit()" style="font-size: 9pt; height: 19; width: 60"> <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='News_Manage.asp'" style="font-size: 9pt; height: 19; width: 60"> </p>
	</form>


</body>
</html>
