<!--#include file="config.asp"-->
<h1>ϵͳ���¼�¼</h1>
<hr>
<%
set rs=server.createobject("adodb.recordset")
sqlcmd = ("select * from upgrade")			'��ѯ���м�¼
rs.Open sqlcmd,Louconn,1,1
a=1
do while not rs.eof
%>
	<div style="line-height: 30px;"><%=a%>��<%=rs("Up_Content")%>��------<%=rs("Up_time")%></div>
<%
a=a+1
rs.movenext
loop
rs.Close
set rs=nothing
%>
	<div id="zhform">
	<form name="form1" method="POST" action="upgrade1.asp">
		<table cellpadding="0" cellspacing="0">
			<tr><td></td></tr>

			<tr>
				<td><textarea id ="Up_Content" name="Up_Content" cols="80" rows="10"></textarea></td>
			</tr>
			<tr><td><input id="upgrade" name=B20 type="submit"  value="д��"></td></tr>
		</table>
	</form>
	</div>
