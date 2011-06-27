<%
Function News_list(NClassID,MaxID,MaxLen)
dim News_list_temp,i
  News_list_temp=""
  News_list_temp=News_list_temp & "<TABLE cellSpacing=""0"" width=""100%"" border=""0"">" &vbcrlf
  set rs_n=server.createobject("adodb.recordset")
  sql_n="select * from newsdata where NClassID="&NClassID&" Order by NewsID Desc"
  rs_n.open sql_n,conn,1,1
  if rs_n.eof and rs_n.bof then
      News_list_temp=News_list_temp&"<tr><td>No Data!</td></tr>" &vbcrlf
  else
  i=0
  do while not rs_n.eof
  topic=gotTopic(rs_n("Title"),MaxLen)
  topic=replace(server.HTMLencode(topic)," ","&nbsp;")
  topic=replace(topic,"'","&nbsp;")
  News_list_temp=News_list_temp & "<TR>" &vbcrlf
  News_list_temp=News_list_temp & "<TD background=images/newsbill22.gif height=23>" &vbcrlf
  News_list_temp=News_list_temp & "<IMG height=10 hspace=5 src=""images/Cnewsbill12.gif"" width=10>"
  News_list_temp=News_list_temp & "<A class=""hui"" href=""NewsDetail.asp?NewsID="&rs_n("NewsID")&""">"&topic&""
  if rs_n("Picture")<>"" and rs_n("IncludePic")=true then
      News_list_temp=News_list_temp & "(ͼ)"
  end if
  News_list_temp=News_list_temp & "</A></TD>" &vbcrlf       
  News_list_temp=News_list_temp & "</TR>" &vbcrlf          
  i=i+1
  if i>=MaxID then exit do
  rs_n.movenext
  loop
  end if
  rs_n.close
  set rs_n=nothing
  News_list_temp=News_list_temp & "</table>"
  response.write News_list_temp
End Function

Function News_left(NClassID,MaxID,MaxLen)
dim News_left_temp,i
  News_left_temp=""
  News_left_temp=News_left_temp & "<TABLE cellSpacing=""0"" width=""100%"" border=""0"">" &vbcrlf
  set rs_n=server.createobject("adodb.recordset")
  sql_n="select * from newsdata where NClassID="&NClassID&" Order by NewsID Desc"
  rs_n.open sql_n,conn,1,1
  if rs_n.eof and rs_n.bof then
      News_left_temp=News_left_temp &"<tr><td>No data!</td></tr>" &VbCrlf
  else
  i=0
  do while not rs_n.eof
  topic=gotTopic(rs_n("Title"),MaxLen)
  topic=replace(server.HTMLencode(topic)," ","&nbsp;")
  topic=replace(topic,"'","&nbsp;")
  News_left_temp=News_left_temp & "<TR>" &vbcrlf
  News_left_temp=News_left_temp & "<TD background=images/newsbill22.gif height=23>" &vbcrlf
  News_left_temp=News_left_temp & "<IMG height=10 hspace=7 src=""images/newsbill4.gif"" width=10 align=absMiddle>"
  News_left_temp=News_left_temp & "<A class=""s"" href=""NewsDetail.asp?NewsID="&rs_n("NewsID")&""">"&topic&""
  if rs_n("Picture")<>"" and rs_n("IncludePic")=true then
      News_left_temp=News_left_temp & "(ͼ)"
  end if
  News_left_temp=News_left_temp & "</A></TD>" &vbcrlf       
  News_left_temp=News_left_temp & "</TR>" &vbcrlf          
  i=i+1
  if i>=MaxID then exit do
  rs_n.movenext
  loop
  end if
  rs_n.close
  set rs_n=nothing
  News_left_temp=News_left_temp & "</table>"
  response.write News_left_temp
End Function

Function News_left2(ss,MaxID,MaxLen,NClassID)
dim News_left_temp,i
  News_left_temp=""
  News_left_temp=News_left_temp & "<TABLE cellSpacing=""0"" width=""100%"" border=""0"">" &vbcrlf
  set rs_n=server.createobject("adodb.recordset")
  sql_n="select * from newsdata where 1=1" 
  if NClassID<>"0" Then
     sql_n=sql_n & " AND NClassID="&NClassID&""
  end if
  if ss="hot" then
     sql_n=sql_n & " Order By Hits Desc"
  elseif ss="Recommend" then
     sql_n=sql_n & " And Recommend=true Order by NewsID Desc"
  else
     sql_n=sql_n & " Order by NewsID Desc"
  end if
  rs_n.open sql_n,conn,1,1
  if rs_n.eof and rs_n.bof then
      News_left_temp=News_left_temp &"<tr><td>No data!</td></tr>" &VbCrlf
  else
  i=0
  do while not rs_n.eof
  topic=gotTopic(rs_n("Title"),MaxLen)
  topic=replace(server.HTMLencode(topic)," ","&nbsp;")
  topic=replace(topic,"'","&nbsp;")
  News_left_temp=News_left_temp & "<TR>" &vbcrlf
  News_left_temp=News_left_temp & "<TD background=images/newsbill22.gif height=23>" &vbcrlf
  News_left_temp=News_left_temp & "<IMG height=10 hspace=7 src=""images/newsbill4.gif"" width=10 align=absMiddle>"
  News_left_temp=News_left_temp & "<A class=""s"" href=""NewsDetail.asp?NewsID="&rs_n("NewsID")&""">"&topic&""
  if rs_n("Picture")<>"" and rs_n("IncludePic")=true then
      News_left_temp=News_left_temp & "(ͼ)"
  end if
  News_left_temp=News_left_temp & "</A></TD>" &vbcrlf       
  News_left_temp=News_left_temp & "</TR>" &vbcrlf          
  i=i+1
  if i>=MaxID then exit do
  rs_n.movenext
  loop
  end if
  rs_n.close
  set rs_n=nothing
  News_left_temp=News_left_temp & "</table>"
  response.write News_left_temp
End Function
%>