<%@language=vbscript codepage=936 %>
<%
'option explicit

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%> 
<!--#include file = "conn.asp"-->
<!--#include file = "function.asp"-->
<!--#include file = "../Login/Check.asp"-->
<!--#include file="../../inc/function.asp"-->
