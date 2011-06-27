<%@ codepage ="936" %>
<!--#include file="../inc/conn.ASP"-->
<%
dim subject,content,F_Company,F_Name,F_Email,F_address,F_Post,F_Tel,F_Fax,T_User,F_User,TF_Time
subject=request("subject")
content=request("content")
F_Company=request("F_Company")
F_Name=request("F_Name")
F_Email=request("F_Email")
F_address=request("F_address")
F_Post=request("F_Post")
F_Tel=request("F_Tel")
F_Fax=request("F_Fax")
T_User=request("T_User")
F_User=request("F_User")
T_Name=request("T_Name")
T_Company=request("T_Company")
T_Email=request("T_Email")
set rs=server.createobject("adodb.recordset")
sql="select * from message"
rs.open sql,conn,1,3
rs.addnew
rs("subject")=subject
rs("content")=content
rs("F_Company")=F_Company
rs("F_Name")=F_Name
rs("F_Email")=F_Email
rs("F_address")=F_address
rs("F_Post")=F_Post
rs("F_Tel")=F_Tel
rs("F_Fax")=F_Fax
rs("T_User")=T_User
rs("F_User")=F_User
rs("T_Name")=T_Name
rs("T_Email")=T_Email
rs("T_Company")=T_Company
rs("TF_Time")=now()
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<div>信息发送成功！<br>

还不是诚信物流网的会员？现在就加入，享受更多会员服务！ 
* 无限制访问产品及供应商的数据库
* 我的办公室轻松管理在线贸易、收发信息、修改资料
* 直接联系商家，发布商业信息，订阅产品速递
* 了解国际贸易法律法规
* 获得海内外贸易发展状况 
还不是诚信物流网的会员？请点此注册，保存您在诚信物流网上将要收发的所有信息。 
已经是诚信物流网的会员？请点此登录，保存您在诚信物流网上将要收发的所有信息。 </div>
