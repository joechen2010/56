<!--#include file="upload.inc"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="check.asp"-->
<html>
<head>
<title>&Icirc;&Auml;&frac14;&thorn;&Eacute;&Iuml;&acute;&laquo;</title>
</head>
<body>
<%
dim fs,file2
set rs=server.createobject("adodb.recordset")
sql="select * from qyml where user='"&session("user")&"'"
rs.open sql,conn,3,2
if not rs.eof then
if rs("cimg")<>"" then
file2=rs("cimg")
set fs=server.CreateObject("scripting.filesystemobject")
file2=server.MapPath(file2)
if fs.FileExists(file2) then
fs.DeleteFile file2,true
end if
rs("cimg")=""
rs.update
end if
rs.close
end if

dim upload,file,formName,formPath,iCount,filename,fileExt,id
set upload=new upload_5xSoft ''�����ϴ�����


 formPath=upload.form("filepath")
 id=upload.form("id")
 ''��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 


iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<100 then
 	response.write "<font size=2>����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>500000 then
 	response.write "<font size=2>�ļ���С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if ( fileEXT<>".gif" and fileEXT<>".jpg") then
 	response.write "<font size=2>�ļ���ʽ���ԡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if 
 
Randomize
	RanNum = Int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&RanNum&FileExt
 
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveAs Server.mappath(filename)   ''�����ļ�
'  response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" �ɹ�!<br>"

'response.end
' response.write "<script>parent.document.forms[1].bigimg.value='"&FileName&"'</script>"
sql="update qyml set cimg='"&filename&"' where user='"&session("user")&"'"
conn.Execute sql
'response.write sql
'response.end
   iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''ɾ���˶���

'session("upface_Bill")="done"

Htmend iCount&" ���ļ��ϴ�����!"

sub HtmEnd(Msg)
 set upload=nothing
%>

 <script language=javascript>alert('ͼƬ�ϴ��ɹ���');window.location.href='cimg_add.asp';</script>
 <%response.end
end sub
%>
</body>
</html>