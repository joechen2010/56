<%
	dim connstr
	dim db,conn
	db="/database/#cn&bearing#.mdb"     '数据库文件的位置
	    'On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
	conn.Open connstr

'释放Rs对象
	Function rsclose
		set rs=nothing
	End Function
	
	 '关闭并释放数据库连接
	Function CloseConn()
		conn.close
		set conn=nothing
	End Function

    Dim UserAgent 
    UserAgent = Trim(Lcase(Request.Servervariables("HTTP_USER_AGENT")))
    If InStr(UserAgent,"teleport") > 0 or InStr(UserAgent,"webzip") > 0 or InStr(UserAgent,"flashget")>0 or InStr(UserAgent,"offline")>0 Then
        Response.Write "请不要采用teleport/Webzip/Flashget/Offline等工具来浏览网站！"
	    Response.End
	End If
	
	dim cf,uf
	cf=conn.execute("select count(id) from qyml where uflag=0")(0)
	uf=conn.execute("select count(id) from qyml where uflag=1")(0)

%>
