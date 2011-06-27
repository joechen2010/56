<%
	dim conn
	dim connstr
	dim db
	db="../../DataBase/#cn&bearing#.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
	conn.Open connstr
	
	'释放Rs对象
	function rsclose
		set rs=nothing
	end function
	
	'关闭并释放数据库连接
	function CloseConn()
		conn.close
		set conn=nothing
	end function
	
dim title
title="天天发物流"
%>
