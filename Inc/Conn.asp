<%
	dim connstr
	dim db,conn
	db="/database/#cn&bearing#.mdb"     '���ݿ��ļ���λ��
	    'On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
	conn.Open connstr

'�ͷ�Rs����
	Function rsclose
		set rs=nothing
	End Function
	
	 '�رղ��ͷ����ݿ�����
	Function CloseConn()
		conn.close
		set conn=nothing
	End Function

    Dim UserAgent 
    UserAgent = Trim(Lcase(Request.Servervariables("HTTP_USER_AGENT")))
    If InStr(UserAgent,"teleport") > 0 or InStr(UserAgent,"webzip") > 0 or InStr(UserAgent,"flashget")>0 or InStr(UserAgent,"offline")>0 Then
        Response.Write "�벻Ҫ����teleport/Webzip/Flashget/Offline�ȹ����������վ��"
	    Response.End
	End If
	
	dim cf,uf
	cf=conn.execute("select count(id) from qyml where uflag=0")(0)
	uf=conn.execute("select count(id) from qyml where uflag=1")(0)

%>
