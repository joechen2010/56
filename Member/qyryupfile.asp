<!--#include file="upload.inc"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<title>&Icirc;&Auml;&frac14;&thorn;&Eacute;&Iuml;&acute;&laquo;</title>
</head>
<body>
<%

dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''建立上传对象

dim zsmc
 formPath="honorUpLoad/"
 zsmc=upload.form("zsmc")
 
 ''在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

'if zsmc="" then
  '    response.write "<script>alert('图片上传成功！');history.back();<'/script>"
'  response.end
 'end if
iCount=0
for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.filesize<10 then
 	response.write "<font size=2>请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>500000 then
 	response.write "<font size=2>文件大小超过了限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if ( fileEXT<>".gif" and fileEXT<>".jpg") then
 	response.write "<font size=2>文件格式不对　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if 
 
Randomize
	RanNum = Int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&RanNum&FileExt
 
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveAs Server.mappath(filename)   ''保存文件
'  response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" 成功!<br>"
'response.end
' response.write "<script>parent.document.forms[1].bigimg.value='"&FileName&"'<'/script>"
sql="insert into qyzs (zsmc,zsurl,gsid,scsj) values ('"&zsmc&"','"&filename&"','"&session("user")&"','"&now()&"')"
conn.Execute sql
'response.write sql
'response.end
   iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''删除此对象

'session("upface_Bill")="done"

Htmend iCount&" 个文件上传结束!"

sub HtmEnd(Msg)
 set upload=nothing
%>

 <script language=javascript>alert('图片上传成功！');window.location.href='qyry.asp';</script>
 <%response.end
end sub
%>
</body>
</html>