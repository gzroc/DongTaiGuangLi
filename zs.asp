<meta http-equiv="refresh" content="20">
<!--#include file="config.asp"-->
<link rel="stylesheet" type="text/css" href="css/NFDT.css">
<div class="main">
<%
set Lours=server.createobject("adodb.recordset")
id=6
%>
<!--#include file="chore.asp"-->
<%
sqlcmd = ("select * from NFDT")			'��ѯ���м�¼
Lours.Open sqlcmd,Louconn,1,1
a=0
dim Count_Rs(4),Count_Wz(4),Count_Gz(4)
dim Fbiaoti
do while not Lours.eof
%> 
	<div class="suo">
		<%
			select case Lours("NFDT_FenSuo")
			case 1
			Fbiaoti="һ����"
			case 2
			Fbiaoti="������"
			case 3
			Fbiaoti="������"
			case 4
			Fbiaoti="�ķ���"
			end select
		%>
		<table>
		<caption><%=Fbiaoti%></caption>
		<tr><th class="th1">��λ</th><th class="th1">����</th><th class="th1">Σ����Ա</th><th class="th1">������ҽ</th><th class="th1">ֵ���쵼</th></tr>
		<tr>
		<td>һ���</td>
		<td><span class="shujv"><%=Lours("NFDT_1DArenshu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaWeiZhong")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaGuangZhou")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaZhiBan")%></span></td>
		</tr>
		<tr>
		<td>�����</td>
		<td><span class="shujv"><%=Lours("NFDT_2DArenshu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_2DaWeiZhong")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_2DaGuangZhou")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_2DaZhiBan")%></span></td>
		</tr>
		<tr>
		<td>�����</td>
		<td><span class="shujv"><%=Lours("NFDT_3DArenshu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_3DaWeiZhong")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_3DaGuangZhou")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_3DaZhiBan")%></span></td>
		</tr>
		<tr>
		<td>
		<%
		if(Lours("NFDT_FenSuo")=4) then
			response.write("��ҽ�ж�")
		else
			response.write("�Ĵ�")
		end if
		%>
		</td>
		<td><span class="shujv"><%=Lours("NFDT_4DArenshu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_4DaWeiZhong")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_4DaGuangZhou")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_4DaZhiBan")%></span></td>
		</tr>
		<tr>
		<td>��������</td>
		<td><span class="shujv"><%=Lours("NFDT_1DArenshu")+Lours("NFDT_2DArenshu")+Lours("NFDT_3DArenshu")+Lours("NFDT_4DArenshu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaWeiZhong")+Lours("NFDT_2DaWeiZhong")+Lours("NFDT_3DaWeiZhong")+Lours("NFDT_4DaWeiZhong")%></span></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaGuangZhou")+Lours("NFDT_2DaGuangZhou")+Lours("NFDT_3DaGuangZhou")+Lours("NFDT_4DaGuangZhou")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_FenSuoZhiBan")%></span></td>
		</tr>
		<%
		if(not Lours("NFDT_DongTai")&"1"="1") then
		%>
		<tr>
		<td>���ն�̬</td><td colspan="4"><marquee direction="left" scrollamount="5" onMouseOver="this.stop();" onMouseOut="this.start();"><span class="dt"><%=Lours("NFDT_DongTai")%></span></marquee></td>
		</tr>
		<%
		end if
		%>
		</table>
		<%Count_Rs(a)=Lours("NFDT_1DArenshu")+Lours("NFDT_2DArenshu")+Lours("NFDT_3DArenshu")+Lours("NFDT_4DArenshu")%>
		<%Count_Wz(a)=Lours("NFDT_1DaWeiZhong")+Lours("NFDT_2DaWeiZhong")+Lours("NFDT_3DaWeiZhong")+Lours("NFDT_4DaWeiZhong")%>
		<%Count_Gz(a)=Lours("NFDT_1DaGuangZhou")+Lours("NFDT_2DaGuangZhou")+Lours("NFDT_3DaGuangZhou")+Lours("NFDT_4DaGuangZhou")%>
		<div class="lastupdate"><span class="baogaoren">����ϱ�ʱ�䣺<%=Lours("NFDT_GenXin")%>&nbsp;&nbsp;&nbsp;&nbsp;�����ˣ�<%=Lours("NFDT_BaoGaoRen")%></span></div>

	</div>
<%
a=a+1
Lours.movenext
loop
%>
<hr>
<div id="count">ȫ����������<span class="zong">
<%
for i=0 to ubound(Count_Rs)
dim count_rs1
count_rs1=Cint(Count_Rs1)+Cint(Count_Rs(i))
next
response.write(count_rs1)
%>��</span>
Σ����Ա��<span class="zong">
<%
for i=0 to ubound(Count_Wz)
dim count_wz1
count_wz1=Cint(Count_wz1)+Cint(Count_wz(i))
next
response.write(count_wz1)
%>��</span>
������ҽ��<span class="zong">
<%
for i=0 to ubound(Count_Gz)
dim count_gz1
count_gz1=Cint(Count_gz1)+Cint(Count_gz(i))
next
response.write(count_gz1)
%>��</span>

</div>

<br>

<%
Lours.Close 
Louconn.Close
set Lours=nothing
set Louconn=nothing
%>
<!--#include file="footer.asp"-->
</div>