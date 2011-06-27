<!--#include FILE="upload.inc"-->
<html>
<head>
<title>&Icirc;&Auml;&frac14;&thorn;&Eacute;&Iuml;&acute;&laquo;</title>
</head>
<body>
<script>
parent.document.forms[0].ToSubmit.disabled=false;
parent.document.forms[0].ToBack.disabled=false;
</script>
<%

dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''建立上传对象


 formPath="smallpic/"
 ''在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 


iCount=0
for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.filesize<100 then
 	response.write "<font size=2>请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>500000 then
 	response.write "<font size=2>文件大小超过了限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if ( fileEXT<>".gif" and fileEXT<>".jpg" and fileEXT<>".avi" and fileEXT<>".swf") then
 	response.write "<font size=2>文件格式不对　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if 
 Randomize
	RanNum = Int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&RanNum&FileExt
 
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveAs Server.mappath(filename)   ''保存文件
'  response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" 成功!<br>"
'response.write filename
'response.end
 response.write "<script>parent.document.forms[0].smallimg.value='"&FileName&"'</script>"
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''删除此对象

session("upface_Bill")="done"

Htmend iCount&" 个文件上传结束!"

sub HtmEnd(Msg)
 set upload=nothing
%>
<SCRIPT LANGUAGE="JavaScript"> alert("文件上传成功,请copy下边的文件连接,以备后用") </script><%
 response.end
end sub


%>
</body>
</html>