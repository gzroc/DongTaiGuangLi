<!--#include file="config.asp"-->
<%

NFDT_SuoZhang=request.form("NFDT_SuoZhang")
NFDT_ChuShi=request.form("NFDT_ChuShi")
NFDT_HuWeiDui=request.form("NFDT_HuWeiDui")
NFDT_YiYuan=request.form("NFDT_YiYuan")
NFDT_ShuiDianBan=request.form("NFDT_ShuiDianBan")
NFDT_SiJi=request.form("NFDT_SiJi")
NFDT_DongTai=replace(request.form("NFDT_DongTai"),vbcrlf,"<br>")
do while instr(NFDT_DongTai,"<br><br>")
	NFDT_DongTai=replace(NFDT_DongTai,"<br><br>","<br>")
'	response.write(NFDT_DongTai)
loop
if(NFDT_DongTai)="<br>" then
	NFDT_DongTai=""
end if
Set Lours=server.createobject("adodb.recordset")
sql="update ZS set NFDT_SuoZhang='"&NFDT_SuoZhang&"',NFDT_ChuShi='"&NFDT_ChuShi&"',NFDT_HuWeiDui='"&NFDT_HuWeiDui&"',NFDT_YiYuan='"&NFDT_YiYuan&"',NFDT_ShuiDianBan='"&NFDT_ShuiDianBan&"',NFDT_SiJi='"&NFDT_SiJi&"',NFDT_DongTai='"&NFDT_DongTai&"',NFDT_GenXin='"&now()&"'"
'response.write sql
'response.end
Lours.Open sql,Louconn,1,3
call writerzzs(NFDT_SuoZhang,NFDT_ChuShi,NFDT_HuWeiDui,NFDT_YiYuan,NFDT_ShuiDianBan,NFDT_SiJi,NFDT_DongTai)
set Lours=nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('信息上报成功');"&"window.location.href = 'zhihuizhongxin.asp'"&" </script>" 
%>