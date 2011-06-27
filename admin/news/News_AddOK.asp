<!--#include file = "../Inc/Syscode.asp"-->
<%
'====================================================
'
'
'=====================================================
Dim i

Dim sNClassID,sClassID, sTitle, sContent,NID
    sNClassID = Request.Form("NClassID")
    sClassID = Request.Form("ClassID")
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
	Sql = "Select * From NewsData"
	Rs.Open Sql,Conn,1,3
	Rs.AddNew
	Rs("NClassID") = sNClassID
	Rs("ClassID")= sClassID
	Rs("Title") = sTitle
	Rs("Content") = sContent
	Rs("Picture") = sPicture
	if sCopyFrom="" then
		Rs("CopyFrom") = ""
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
	NID=Rs("NewsID")
	Rs.Close
	Conn.Execute("UPDATE NewsData SET Updown="&NID&" WHERE NewsID="&NID&"")
       response.Write "<script language='javascript'>alert('数据添加成功！');window.location.href='news_add.asp';</script>"
%>

