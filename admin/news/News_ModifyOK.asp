<!--#include file = "../Inc/Syscode.asp"-->
<%
'====================================================
'
'
'=====================================================
Dim i

Dim sNewsID,sClassID,sNClassID, sTitle, sContent
    sNewsID = request("NewsID")
    sClassID = Request.Form("ClassID")
    sNClassID = Request.Form("NClassID")
	sTitle = Request.Form("Title")
	
	sContent = ""
	For i = 1 To Request.Form("Content").Count
		sContent = sContent & Request.Form("Content")(i)
	Next
	
    Dim  sPicture, sCopyFrom, sIncludePic, sRecommend, sEditor
    sPicture = Request.Form("Picture")
    sCopyFrom = Request.Form("CopyFrom")
	sIncludePic = Request.Form("IncludePic")
	sRecommend = Request.Form("Recommend")
    sEditor = Session("AdminName")
	
	Dim NewsID,Rs,Sql
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Sql = "Select * From NewsData where NewsID="&sNewsID&""
	Rs.Open Sql,Conn,1,3
	Rs("ClassID") = sClassID
	Rs("NClassID") = sNClassID
	Rs("Title") = sTitle
	Rs("Content") = sContent
	Rs("Picture") = sPicture
	if sCopyFrom="" then
		Rs("CopyFrom") = "天天发物流"
	else
	    Rs("CopyFrom") = sCopyFrom
	end if
	if sIncludePic="yes" then
		Rs("IncludePic") = true
	else
	   Rs("IncludePic") = false
	end if
	Rs("Editor") = sEditor
	if sRecommend="yes" then
		Rs("Recommend") = true
	else
	    Rs("Recommend") = false
	end if
	Rs.Update
	Rs.Close
    response.Write "<script language='javascript'>alert('新闻修改成功！');window.location.href='News_Manage.asp';</script>"
%>

