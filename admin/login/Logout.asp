<%
Response.Buffer=true     
Response.Expires=0         
Response.CacheControl="no-cache"

session("user")=""
session.abandon
response.write "<script>top.location.href='../index.asp';</script>"
%>