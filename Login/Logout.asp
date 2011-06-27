<%@ codepage ="936" %>
<%
Response.Buffer=true     
Response.Expires=0         
Response.CacheControl="no-cache"

session("user")=""
session.abandon
response.redirect "../index.asp"
%>