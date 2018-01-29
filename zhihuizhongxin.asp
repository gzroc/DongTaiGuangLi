<meta http-equiv="refresh" content="180">
<!--#include file="config.asp"-->
<link type="text/css" rel="stylesheet" href="css/NFDT1.css"/> 
<head>
<link rel="shortcut icon" href="image/nf.ico">
</head>
<div class="main">
<%
set rs1=server.createobject("adodb.recordset")
str = ("select  * from ZS")			'查询id(ID号)这条记录
rs1.Open str,Louconn,1,1

set rs=server.createobject("adodb.recordset")
sqlcmd = ("select * from NFDT")			'查询所有记录
rs.Open sqlcmd,Louconn,1,1
a=0
dim Count_Rs(4),Count_Wz(4),Count_Gz(4),Count_Dd(4),Count_Mj(4)
do while not rs.eof
%>
		<%Count_Rs(a)=rs("NFDT_1DArenshu")+rs("NFDT_2DArenshu")+rs("NFDT_3DArenshu")+rs("NFDT_4DArenshu")%>
		<%Count_Wz(a)=rs("NFDT_1DaWeiZhong")+rs("NFDT_2DaWeiZhong")+rs("NFDT_3DaWeiZhong")+rs("NFDT_4DaWeiZhong")%>
		<%Count_Gz(a)=rs("NFDT_1DaGuangZhou")+rs("NFDT_2DaGuangZhou")+rs("NFDT_3DaGuangZhou")+rs("NFDT_4DaGuangZhou")%>
		<%Count_Dd(a)=rs("NFDT_1DaDanDu")+rs("NFDT_2DaDanDu")+rs("NFDT_3DaDanDu")+rs("NFDT_4DaDanDu")%>
		<%Count_Mj(a)=rs("NFDT_1DaMinJing")+rs("NFDT_2DaMinJing")+rs("NFDT_3DaMinJing")+rs("NFDT_4DaMinJing")%>
<%
a=a+1
rs.movenext
loop
rs.Close
set rs=nothing
%>

<%
set Lours=server.createobject("adodb.recordset")
id=5
'分所表格数量'
dim counts
%>
<!--#include file="chore.asp"-->
<div id="count">全所总人数：<span class="zong">
<%
for i=0 to ubound(Count_Rs)
dim count_rs1
count_rs1=Cint(Count_Rs1)+Cint(Count_Rs(i))
next
response.write(count_rs1)
%>人</span>
危重人员：<span class="zong">
<%
for i=0 to ubound(Count_Wz)
dim count_wz1
count_wz1=Cint(Count_wz1)+Cint(Count_wz(i))
next
response.write(count_wz1)
%>人</span>
单独管理：<span class="zong">
<%
for i=0 to ubound(Count_Dd)
dim count_dd1
count_dd1=Cint(Count_dd1)+Cint(Count_dd(i))
next
response.write(count_dd1)
%>人</span>
所外留医：<span class="zong">
<%
for i=0 to ubound(Count_Gz)
dim count_gz1
count_gz1=Cint(Count_gz1)+Cint(Count_gz(i))
next
response.write(count_gz1)
%>人</span>
在队民警：<span class="zong">
<%
for i=0 to ubound(Count_Mj)
dim count_mj1
count_mj1=Cint(Count_mj1)+Cint(Count_mj(i))
next
response.write(count_mj1)
%>人</span>
<span class="noprint"><input type="button" id="printbutton" name="printbutton" onclick="javascript:window.print()" value="打印" height="50"></span>
</div>
<%
Lours.Open sqlcmd,Louconn,1,1
counts=Lours.recordcount
'response.write(counts)
dim Fbiaoti 
%>
<!--下面是总所值班表单-->
<div class="noprint">
	<div id="zhform">
	<form name="form1" method="POST" action="update1.asp">
		<table cellpadding="0" cellspacing="0">
			<tr>
				<td>所长</td>
				<td><input type="text" id="NFDT_SuoZhang" name="NFDT_SuoZhang" class="button3" value="<%=rs1("NFDT_SuoZhang")%>"></td>
				<td>处室</td>
				<td><input type="text" id="NFDT_ChuShi" name="NFDT_ChuShi" class="button4" value="<%=rs1("NFDT_ChuShi")%>"></td>
				<td>护卫队</td>
				<td><input type="text" id="NFDT_HuWeiDui" name="NFDT_HuWeiDui" class="button3" value="<%=rs1("NFDT_HuWeiDui")%>"></td>
				<td>医院</td>
				<td><input type="text" id="NFDT_YiYuan" name="NFDT_YiYuan" class="button3" value="<%=rs1("NFDT_YiYuan")%>"></td>
				<td>司机</td>
				<td><input type="text" id="NFDT_SiJi" name="NFDT_SiJi" class="button4" value="<%=rs1("NFDT_SiJi")%>"></td>
				<td>水电班</td>
				<td><input type="text" id="NFDT_ShuiDianBan" name="NFDT_ShuiDianBan" class="button3" value="<%=rs1("NFDT_ShuiDianBan")%>"></td>
			</tr>
			<tr>
				<td>当日动态</td>
				<td colspan="10"><textarea id="textare1" name="NFDT_DongTai"><%=replace(rs1("NFDT_DongTai"),"<br>",vbcrlf)%></textarea></td>
				<td ><input id="update1" name=B20 type="button"  onclick="sub3()" value="上报"></td>
			</tr>
		</table>
	</form>
	</div>
</div>
<!--上面是总所值班表单-->
		<table cellpadding="0" cellspacing="0" border="2">
		<tr><th>单位</th><th>大队</th><th>人数</th><th>危重人员</th><th>单独管理</th><th>所外留医</th><th>值班领导</th><th>在队民警</th></tr>
<%
do while not Lours.eof
%> 
		<%
			select case Lours("NFDT_FenSuo")
			case 1
			Fbiaoti="一分所"
			case 2
			Fbiaoti="二分所"
			case 3
			Fbiaoti="三分所"
			case 4
			Fbiaoti="四分所"
			end select
		%>

		<tr><td rowspan="6"><%=Fbiaoti%></td>
		<td>一大队</td>
		<td><span class="shujv"><%=Lours("NFDT_1DArenshu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaWeiZhong")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaDanDu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaGuangZhou")%></span></td>
		<td><span class="shujv"><%=dantian(Lours("NFDT_1DaZhiBan"),Lours("NFDT_GenXin"))%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaMinJing")%></span></td>
		</tr>
		<tr>
		<td>二大队</td>
		<td><span class="shujv"><%=Lours("NFDT_2DArenshu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_2DaWeiZhong")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_2DaDanDu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_2DaGuangZhou")%></span></td>
		<td><span class="shujv"><%=dantian(Lours("NFDT_2DaZhiBan"),Lours("NFDT_GenXin"))%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_2DaMinJing")%></span></td>
		</tr>
		<tr>
		<td>三大队</td>
		<td><span class="shujv"><%=Lours("NFDT_3DArenshu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_3DaWeiZhong")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_3DaDanDu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_3DaGuangZhou")%></span></td>
		<td><span class="shujv"><%=dantian(Lours("NFDT_3DaZhiBan"),Lours("NFDT_GenXin"))%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_3DaMinJing")%></span></td>
		</tr>
		<tr>
		<td>
		<%
		if(Lours("NFDT_FenSuo")=4) then
			response.write("留医中队")
		else
			response.write("四大队")
		end if
		%>
		</td>
		<td><span class="shujv"><%=Lours("NFDT_4DArenshu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_4DaWeiZhong")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_4DaDanDu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_4DaGuangZhou")%></span></td>
		<td><span class="shujv"><%=dantian(Lours("NFDT_4DaZhiBan"),Lours("NFDT_GenXin"))%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_4DaMinJing")%></span></td>
		</tr>
		<tr>
		<td>分所机关</td>
		<td><span class="shujv"><%=Lours("NFDT_1DArenshu")+Lours("NFDT_2DArenshu")+Lours("NFDT_3DArenshu")+Lours("NFDT_4DArenshu")%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaWeiZhong")+Lours("NFDT_2DaWeiZhong")+Lours("NFDT_3DaWeiZhong")+Lours("NFDT_4DaWeiZhong")%></span></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaDanDu")+Lours("NFDT_2DaDanDu")+Lours("NFDT_3DaDanDu")+Lours("NFDT_4DaDanDu")%></span></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaGuangZhou")+Lours("NFDT_2DaGuangZhou")+Lours("NFDT_3DaGuangZhou")+Lours("NFDT_4DaGuangZhou")%></span></td>
		<td><span class="shujv"><%=dantian(Lours("NFDT_FenSuoZhiBan"),Lours("NFDT_GenXin"))%></span></td>
		<td><span class="shujv"><%=Lours("NFDT_1DaMinJing")+Lours("NFDT_2DaMinJing")+Lours("NFDT_3DaMinJing")+Lours("NFDT_4DaMinJing")%></span></td>
		</tr>
		<tr>
		<td >当日动态</td><td colspan="6"><!--marquee direction="left" scrollamount="5" onMouseOver="this.stop();" onMouseOut="this.start();"--><div class="dt"><%=dantian(Lours("NFDT_DongTai"),Lours("NFDT_GenXin"))%></div><!--/marquee-->	
		<span class="baogaoren">最后报告时间：<%=formatdatetime(Lours("NFDT_GenXin"),1)%><%=formatdatetime(Lours("NFDT_GenXin"),4)%>&nbsp;&nbsp;报告人：<%=Lours("NFDT_BaoGaoRen")%></span></td></tr>
		<!--div class="lastupdate"><%=Fbiaoti%>最后上报时间：<%=Lours("NFDT_GenXin")%><span class="baogaoren">报告人：<%=Lours("NFDT_BaoGaoRen")%></span></div-->

<%
a=a+1
Lours.movenext
loop
%>
		</table>

<hr>

<%
Lours.Close 
Louconn.Close
set rs1=nothing
set Lours=nothing
set Louconn=nothing
%>
<!--#include file="footer.asp"-->
</div>