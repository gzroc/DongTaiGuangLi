<!--#include file="config.asp"-->
<link rel="stylesheet" type="text/css" href="css/NFDT.css">
<%
set rs1=server.createobject("adodb.recordset")
str = ("select * from ZS")			'��ѯ����ֵ���¼
rs1.Open str,Louconn,1,1
dim zbsz,zbcs,hwd,yy,dt,sj
		zbsz=rs1("NFDT_SuoZhang")
		zbcs=rs1("NFDT_ChuShi")
		hwd=rs1("NFDT_HuWeiDui")
		yy=rs1("NFDT_YiYuan")
		sdb=rs1("NFDT_ShuiDianBan")
		sji=rs1("NFDT_SiJi")
		dt=rs1("NFDT_DongTai")
		sj=formatdatetime(rs1("NFDT_GenXin"),2)
rs1.Close
set rs1=nothing
%>
<%
set Lours=server.createobject("adodb.recordset")
id=1
sqlcmd = ("select  * from NFDT where NFDT_FenSuo="&id)			'��ѯid(ID��)������¼
Lours.Open sqlcmd,Louconn,1,1
%> 
<div class="main">
<!--#include file="chore.asp"-->
<table>		
	<tr><td rowspan="2">ʱ��</td><td rowspan="2"><span class="zszbzf"><%=sj%></span></td><td>����</td><td><span class="zszbzf"><%=zbsz%></span></td><td>����</td><td colspan="3"><span class="zszbzf"><%=zbcs%></span></td><td>ҽԺ</td><td><span class="zszbzf"><%=yy%></span></td></tr>
	<tr><td>������</td><td><span class="zszbzf"><%=hwd%></span></td><td>˾��</td><td colspan="3"><span class="zszbzf"><%=sji%></span></td><td>ˮ���</td><td><span class="zszbzf"><%=sdb%></span></td></tr>
</table>
</table><table cellspacing="5" cellpadding="5" id="table1">
<form name="form1" method="POST" action="update.asp?id=<%=Lours("id")%>">
<tr><th>������λ</th><th>����</th><th>Σ����Ա</th><th>��������</th><th>������ҽ</th><th>ֵ���쵼</th><th>�ڶ���</th></tr>
<tr> 
<td>һ���</td>
<td><input type="text" id="NFDT_1DArenshu" name="NFDT_1DArenshu" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_1DArenshu")%>" ></td>
<td><input type="text" id="NFDT_1DaWeiZhong" name="NFDT_1DaWeiZhong" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_1DaWeiZhong")%>"></td>
<td><input type="text" id="NFDT_1DaDanDu" name="NFDT_1DaDanDu" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_1DaDanDu")%>"></td>
<td><input type="text" id="NFDT_1DaGuangZhou" name="NFDT_1DaGuangZhou" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_1DaGuangZhou")%>"></td>
<td><input type="text" id="NFDT_1DaZhiBan" name="NFDT_1DaZhiBan" class="button1" value="<%=Lours("NFDT_1DaZhiBan")%>"></td>
<td><input type="text" id="NFDT_1DaMinJing" name="NFDT_1DaMinJing" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_1DaMinJing")%>"></td>

</tr>
<tr> 
<td>�����</td>
<td><input type="text" id="NFDT_2DArenshu" name="NFDT_2DArenshu" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_2DArenshu")%>"></td>
<td><input id="NFDT_2DaWeiZhong" name="NFDT_2DaWeiZhong" maxlength="100" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_2DaWeiZhong")%>"></td>
<td><input id="NFDT_2DaDanDu" name="NFDT_2DaDanDu" maxlength="100" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_2DaDanDu")%>"></td>
<td><input type="text" id="NFDT_2DaGuangZhou" name="NFDT_2DaGuangZhou" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_2DaGuangZhou")%>"></td>
<td><input type="text" id="NFDT_2DaZhiBan" name="NFDT_2DaZhiBan" class="button1" value="<%=Lours("NFDT_2DaZhiBan")%>"></td>
<td><input type="text" id="NFDT_2DaMinJing" name="NFDT_2DaMinJing" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_2DaMinJing")%>"></td>

</tr>
<tr> 
<td>�����</td>
<td><input type="text" id="NFDT_3DArenshu" name="NFDT_3DArenshu" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_3DArenshu")%>"></td>
<td><input type="text" id="NFDT_3DaWeiZhong" name="NFDT_3DaWeiZhong" maxlength="30" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_3DaWeiZhong")%>"></td>
<td><input type="text" id="NFDT_3DaDanDu" name="NFDT_3DaDanDu" maxlength="30" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_3DaDanDu")%>"></td>
<td><input type="text" id="NFDT_3DaGuangZhou" name="NFDT_3DaGuangZhou" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_3DaGuangZhou")%>"></td>
<td><input type="text" id="NFDT_3DaZhiBan" name="NFDT_3DaZhiBan" class="button1" value="<%=Lours("NFDT_3DaZhiBan")%>"></td>
<td><input type="text" id="NFDT_3DaMinJing" name="NFDT_3DaMinJing" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_3DaMinJing")%>"></td>

</tr>
<tr> 
<td>
<%
if(id=4) then
response.write("��ҽ�ж�")
else
response.write("�Ĵ��")
end if
%></td>
<td><input type="text" id="NFDT_4DArenshu" name="NFDT_4DArenshu" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_4DArenshu")%>"></td>
<td><input type="text" id="NFDT_4DaWeiZhong" name="NFDT_4DaWeiZhong" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_4DaWeiZhong")%>"></td>
<td><input type="text" id="NFDT_4DaDanDu" name="NFDT_4DaDanDu" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_4DaDanDu")%>"></td>
<td><input type="text" id="NFDT_4DaGuangZhou" name="NFDT_4DaGuangZhou" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_4DaGuangZhou")%>"></td>
<td><input type="text" id="NFDT_4DaZhiBan" name="NFDT_4DaZhiBan" class="button1" value="<%=Lours("NFDT_4DaZhiBan")%>"></td>
<td><input type="text" id="NFDT_4DaMinJing" name="NFDT_4DaMinJing" class="button1" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%=Lours("NFDT_4DaMinJing")%>"></td>
</tr>
<tr>
<td>���ն�̬</td>
<td colspan="6">
<textarea id="textare" name="NFDT_DongTai"><%=replace(Lours("NFDT_DongTai"),"<br>",vbcrlf)%>
</textarea><br><div class="sili">��дʾ����8:00��X�����XX��ѧԱ�������侯ҽԺ������</div>
</td>
</tr>
<tr> 
<td>����ֵ��</td>
<td > 
<input type="text" id="NFDT_FenSuoZhiBan" name="NFDT_FenSuoZhiBan" class="button2" value="<%=Lours("NFDT_FenSuoZhiBan")%>">
</td>
<td colspan="5"><span id="countrs">������:<font color="blue"><%=Lours("NFDT_1DArenshu")+Lours("NFDT_2DArenshu")+Lours("NFDT_3DArenshu")+Lours("NFDT_4DArenshu")%>&nbsp;&nbsp;&nbsp;&nbsp;</font>Σ������:<font color="blue"><%=Lours("NFDT_1DaWeiZhong")+Lours("NFDT_2DaWeiZhong")+Lours("NFDT_3DaWeiZhong")+Lours("NFDT_4DaWeiZhong")%>&nbsp;&nbsp;&nbsp;&nbsp;</font>��������:<font color="blue"><%=Lours("NFDT_1DaDanDu")+Lours("NFDT_2DaDanDu")+Lours("NFDT_3DaDanDu")+Lours("NFDT_4DaDanDu")%>&nbsp;&nbsp;&nbsp;&nbsp;</font>������ҽ:<font color="blue"><%=Lours("NFDT_1DaGuangZhou")+Lours("NFDT_2DaGuangZhou")+Lours("NFDT_3DaGuangZhou")+Lours("NFDT_4DaGuangZhou")%>&nbsp;&nbsp;&nbsp;&nbsp;</font>�ڶ���:<font color="blue"><%=Lours("NFDT_1DaMinJing")+Lours("NFDT_2DaMinJing")+Lours("NFDT_3DaMinJing")+Lours("NFDT_4DaMinJing")%>&nbsp;&nbsp;&nbsp;&nbsp;</font></span></td>
</tr>
<tr> 
<td>������</td>
<td><input type="text" name="NFDT_BaoGaoRen" id="NFDT_BaoGaoRen" class="button1" value="<%=Lours("NFDT_BaoGaoRen")%>" ></td>
<td colspan="4" ><span class="NoUpdate"><%=NoUpdate(Lours("NFDT_GenXin"))%></span><span class="lastupdate">������ʱ��:<%=FormatDateTime(Lours("NFDT_GenXin"),0)%></span></td><td>���������������� 
<input id="update" name=B12 type="button"  onclick="sub()" value="�ϱ�">
<input type="hidden" name="id" value="<%=id%>">
</td>
</tr>
</form>
</table>
<br>
<!--#include file="footer.asp"-->

<br>

<%
Lours.Close 
Louconn.Close
set Lours=nothing
set Louconn=nothing
%>
</div>