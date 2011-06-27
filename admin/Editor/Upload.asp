<!--#include file="Include/Startup.asp"-->
<!--#include file="Include/upfile_class.asp"-->
<%
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'★                                                                  ★
'☆                eWebEditor - eWebSoft在线文本编辑器               ☆
'★                                                                  ★
'☆  版权所有: eWebSoft.com                                          ☆
'★                                                                  ★
'☆  程序制作: eWeb开发团队                                          ☆
'★            email:webmaster@webasp.net                            ★
'☆            QQ:589808                                             ☆
'★                                                                  ★
'☆  相关网址: [产品介绍]http://www.eWebSoft.com/Product/eWebEditor/ ☆
'★            [支持论坛]http://bbs.eWebSoft.com/                    ★
'☆                                                                  ☆
'★  主页地址: http://www.eWebSoft.com/   eWebSoft团队及产品         ★
'☆            http://www.webasp.net/     WEB技术及应用资源网站      ☆
'★            http://bbs.webasp.net/     WEB技术交流论坛            ★
'★                                                                  ★
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
%>

<%
Server.ScriptTimeOut = 1800
Dim sType, sStyleName
Dim sAllowExt, nAllowSize, sUploadDir

Call InitUpload()		' 初始化上传变量
Call DBConnEnd()		' 断开数据库连接


Dim sAction
sAction = UCase(Trim(Request.QueryString("action")))

Call ShowForm()			' 显示上传表单
If sAction = "SAVE" Then
	Call DoSave()		' 存文件
End If



Sub ShowForm() 
%>
<HTML>
<HEAD>
<TITLE>文件上传</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
body, a, table, div, span, td, th, input, select{font:9pt;font-family: "宋体", Verdana, Arial, Helvetica, sans-serif;}
body {padding:0px;margin:0px}
</style>

<script language="JavaScript" src="Dialog/dialog.js"></script>

</head>
<body bgcolor=menu>

<form action="../htmledit/?action=save&type=<%=sType%>&style=<%=sStyleName%>" method=post name=myform enctype="multipart/form-data">
<input type=file name=uploadfile size=1 style="width:100%">
</form>

<script language=javascript>

var sAllowExt = "<%=sAllowExt%>";
// 检测上传表单
function CheckUploadForm() {
	if (!IsExt(document.myform.uploadfile.value,sAllowExt)){
		parent.UploadError("提示：\n\n请选择一个有效的文件，\n支持的格式有（"+sAllowExt+"）！");
		return false;
	}
	return true
}

// 提交事件加入检测表单
var oForm = document.myform ;
oForm.attachEvent("onsubmit", CheckUploadForm) ;
if (! oForm.submitUpload) oForm.submitUpload = new Array() ;
oForm.submitUpload[oForm.submitUpload.length] = CheckUploadForm ;
if (! oForm.originalSubmit) {
	oForm.originalSubmit = oForm.submit ;
	oForm.submit = function() {
		if (this.submitUpload) {
			for (var i = 0 ; i < this.submitUpload.length ; i++) {
				this.submitUpload[i]() ;
			}
		}
		this.originalSubmit() ;
	}
}

// 上传表单已装入完成
try {
	parent.UploadLoaded();
}
catch(e){
}

</script>

</body>
</html>
<% 
End Sub 

' 保存操作
Sub DoSave()
	Dim oUpload, oFile, sFileExt, sFileName
	' 建立上传对象
	Set oUpload = New upfile_class
	' 取得上传数据,限制最大上传
	oUpload.GetData(nAllowSize*1024)
	If oUpload.Err > 0 Then
		Select Case oUpload.Err
		Case 1
			Call OutScript("parent.UploadError('请选择有效的上传文件！')")
		Case 2
			Call OutScript("parent.UploadError('你上传的文件总大小超出了最大限制（" & nAllowSize & "KB）！')")
		End Select
		Response.End
	End If

	Set oFile = oUpload.File("uploadfile")
	sFileExt = UCase(oFile.FileExt)
	Call CheckValidExt(sFileExt)

	Dim sRnd
	Randomize
	sRnd = Int(900 * Rnd) + 100
	sFileName = year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & sRnd & "." & sFileExt
	oFile.SaveToFile Server.Mappath(sUploadDir & "/"& sFileName)
	
	Set oFile = Nothing
	Set oUpload = Nothing

	Call OutScript("parent.UploadSaved('" & sFileName & "')")

End Sub

' 输出客户端脚本
Sub OutScript(str)
	Response.Write "<script language=javascript>" & str & ";history.back()</script>"
End Sub

' 检测扩展名的有效性
Sub CheckValidExt(sExt)
	Dim b, i, aExt
	b = False
	aExt = Split(sAllowExt, "|")
	For i = 0 To UBound(aExt)
		If UCase(aExt(i)) = UCase(sExt) Then
			b = True
			Exit For
		End If
	Next
	If b = False Then
		OutScript("parent.UploadError('提示：\n\n请选择一个有效的文件，\n支持的格式有（"+sAllowExt+"）！')")
		Response.End
	End If
End Sub


' 初始化上传限制数据
Sub InitUpload()
	sType = UCase(Trim(Request.QueryString("type")))
	sStyleName = Trim(Request.QueryString("style"))
	sSql = "select * from ewebeditor_style where s_name='" & sStyleName & "'"
	oRs.Open sSql, oConn, 0, 1
	If Not oRs.Eof Then
		sUploadDir = oRs("S_UploadDir")
		Select Case sType
		Case "FILE"
			sAllowExt = oRs("S_FileExt")
			nAllowSize = oRs("S_FileSize")
		Case "MEDIA"
			sAllowExt = oRs("S_MediaExt")
			nAllowSize = oRs("S_MediaSize")
		Case "FLASH"
			sAllowExt = oRs("S_FlashExt")
			nAllowSize = oRs("S_FlashSize")
		Case Else
			sAllowExt = oRs("S_ImageExt")
			nAllowSize = oRs("S_ImageSize")
		End Select
	Else
		OutScript("parent.UploadError('无效的样式ID号，请通过页面上的链接进行操作！')")
	End If
	oRs.Close
	' 任何情况下都不允许上传asp脚本文件
	sAllowExt = Replace(UCase(sAllowExt), "ASP", "")
End Sub
%>