<!--#include file="config.asp"-->
<%
up_content1=request.form("Up_Content")
set rs1=server.createobject("adodb.recordset")
str = ("insert into upgrade(up_content)values('"&up_content1&"')")			'��ѯid(ID��)������¼
'response.write(str)
Louconn.execute str


rs1.close
Louconn.close
set rs1=nothing
set Louconn=nothing
response.write"<script language='javascript'>alert('�����ύ�ɹ�');</script>"
response.redirect "upgrade.asp" 
%>