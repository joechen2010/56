<!--#include file="../inc/syscode.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<body>
<%Set rs = Server.CreateObject("ADODB.Recordset")    
	sele="select * from shejishi order by id desc"
	rs.Open sele,conn,3,2%>
	<%=rs("mome")%>
</body>
</html>
