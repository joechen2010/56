<%@language=vbscript codepage=936 %>
<%
'option explicit

'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%> 
<!--#include file = "conn.asp"-->
<!--#include file = "function.asp"-->
<!--#include file = "../Login/Check.asp"-->
<!--#include file="../../inc/function.asp"-->
