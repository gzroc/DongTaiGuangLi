<%
on error resume next

dim Louconn,connstr,db
db="NFDT.mdb" 
Set Louconn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
Louconn.Open connstr

If Err then
Err.Clear
Set Louconn = Nothing
Response.Write "ϵͳ������......���Ժ����ԣ���"
Response.End
End If
%>
<!--#include file="func.asp"-->
