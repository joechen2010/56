<%
	dim conn
	dim connstr
	dim db
	db="../../DataBase/#cn&bearing#.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
	conn.Open connstr
	
	'�ͷ�Rs����
	function rsclose
		set rs=nothing
	end function
	
	'�رղ��ͷ����ݿ�����
	function CloseConn()
		conn.close
		set conn=nothing
	end function
	
dim title
title="���췢����"
%>
