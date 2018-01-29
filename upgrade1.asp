<!--#include file="config.asp"-->
<%
up_content1=request.form("Up_Content")
set rs1=server.createobject("adodb.recordset")
str = ("insert into upgrade(up_content)values('"&up_content1&"')")			'查询id(ID号)这条记录
'response.write(str)
Louconn.execute str


rs1.close
Louconn.close
set rs1=nothing
set Louconn=nothing
response.write"<script language='javascript'>alert('数据提交成功');</script>"
response.redirect "upgrade.asp" 
%>