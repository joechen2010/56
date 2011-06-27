<!--#include file = "../Inc/Syscode.asp"-->
<%
'====================================================
'
'
'=====================================================
Dim i
	sContent = ""
	For i = 1 To Request.Form("Content").Count
		sContent = sContent & Request.Form("Content")(i)
	Next

	Dim NewsID,Rs,Sql
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Sql = "Select * From gonggao where id="&request("id")&""
	Rs.Open Sql,Conn,1,3
	Rs("gonggao") = sContent
	rs("bt")=request("bt")
	rs("addtime")=now()
	Rs.Update

       response.Write "<script language='javascript'>alert('数据修改成功！');window.location.href='gong_Manage.asp';</script>"
%>

