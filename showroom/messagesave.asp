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
<div>��Ϣ���ͳɹ���<br>

�����ǳ����������Ļ�Ա�����ھͼ��룬���ܸ����Ա���� 
* �����Ʒ��ʲ�Ʒ����Ӧ�̵����ݿ�
* �ҵİ칫�����ɹ�������ó�ס��շ���Ϣ���޸�����
* ֱ����ϵ�̼ң�������ҵ��Ϣ�����Ĳ�Ʒ�ٵ�
* �˽����ó�׷��ɷ���
* ��ú�����ó�׷�չ״�� 
�����ǳ����������Ļ�Ա������ע�ᣬ�������ڳ����������Ͻ�Ҫ�շ���������Ϣ�� 
�Ѿ��ǳ����������Ļ�Ա�����˵�¼���������ڳ����������Ͻ�Ҫ�շ���������Ϣ�� </div>
