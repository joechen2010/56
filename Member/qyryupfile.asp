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
set upload=new upload_5xSoft ''�����ϴ�����

dim zsmc
 formPath="honorUpLoad/"
 zsmc=upload.form("zsmc")
 
 ''��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

'if zsmc="" then
  '    response.write "<script>alert('ͼƬ�ϴ��ɹ���');history.back();<'/script>"
'  response.end
 'end if
iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<10 then
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
' response.write "<script>parent.document.forms[1].bigimg.value='"&FileName&"'<'/script>"
sql="insert into qyzs (zsmc,zsurl,gsid,scsj) values ('"&zsmc&"','"&filename&"','"&session("user")&"','"&now()&"')"
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

 <script language=javascript>alert('ͼƬ�ϴ��ɹ���');window.location.href='qyry.asp';</script>
 <%response.end
end sub
%>
</body>
</html>