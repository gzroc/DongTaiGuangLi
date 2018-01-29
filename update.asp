<!--#include file="config.asp"-->
<%
dim IID
IID=cint(request.form("id"))

'if request("LHB_IP")="" then
''	response.write "<script language=JavaScript>" & chr(13) & "alert('请正确输入IP地址！');" & "history.back()" & "</script>" 
''	Response.End
'end if
NFDT_1DArenshu=request.form("NFDT_1DArenshu")
NFDT_2DArenshu=request.form("NFDT_2DArenshu")
NFDT_3DArenshu=request.form("NFDT_3DArenshu")
NFDT_4DArenshu=request.form("NFDT_4DArenshu")
NFDT_1DaWeiZhong=request.form("NFDT_1DaWeiZhong")
NFDT_2DaWeiZhong=request.form("NFDT_2DaWeiZhong")
NFDT_3DaWeiZhong=request.form("NFDT_3DaWeiZhong")
NFDT_4DaWeiZhong=request.form("NFDT_4DaWeiZhong")
NFDT_1DaGuangZhou=request.form("NFDT_1DaGuangZhou")
NFDT_2DaGuangZhou=request.form("NFDT_2DaGuangZhou")
NFDT_3DaGuangZhou=request.form("NFDT_3DaGuangZhou")
NFDT_4DaGuangZhou=request.form("NFDT_4DaGuangZhou")
NFDT_1DaZhiBan=request.form("NFDT_1DaZhiBan")
NFDT_2DaZhiBan=request.form("NFDT_2DaZhiBan")
NFDT_3DaZhiBan=request.form("NFDT_3DaZhiBan")
NFDT_4DaZhiBan=request.form("NFDT_4DaZhiBan")
NFDT_1DaDanDu=request.form("NFDT_1DaDanDu")
NFDT_2DaDanDu=request.form("NFDT_2DaDanDu")
NFDT_3DaDanDu=request.form("NFDT_3DaDanDu")
NFDT_4DaDanDu=request.form("NFDT_4DaDanDu")
NFDT_1DaMinJing=request.form("NFDT_1DaMinJing")
NFDT_2DaMinJing=request.form("NFDT_2DaMinJing")
NFDT_3DaMinJing=request.form("NFDT_3DaMinJing")
NFDT_4DaMinJing=request.form("NFDT_4DaMinJing")
'response.write(NFDT_1DaDanDu)
NFDT_FenSuoZhiBan=request.form("NFDT_FenSuoZhiBan")
NFDT_DongTai=replace(request.form("NFDT_DongTai"),vbcrlf,"<br>")
do while instr(NFDT_DongTai,"<br><br>")
	NFDT_DongTai=replace(NFDT_DongTai,"<br><br>","<br>")
'	response.write(NFDT_DongTai)
loop
if(NFDT_DongTai)="<br>" then
	NFDT_DongTai=""
end if
NFDT_BaoGaoRen=request.form("NFDT_BaoGaoRen")
Set Lours=server.createobject("adodb.recordset")
sql="update NFDT set NFDT_1DArenshu="&NFDT_1DArenshu&",NFDT_2DArenshu="&NFDT_2DArenshu&",NFDT_3DArenshu="&NFDT_3DArenshu&",NFDT_4DArenshu="&NFDT_4DArenshu&",NFDT_1DaWeiZhong="&NFDT_1DaWeiZhong&",NFDT_2DaWeiZhong="&NFDT_2DaWeiZhong&",NFDT_3DaWeiZhong="&NFDT_3DaWeiZhong&",NFDT_4DaWeiZhong="&NFDT_4DaWeiZhong&",NFDT_1DaGuangZhou="&NFDT_1DaGuangZhou&",NFDT_2DaGuangZhou="&NFDT_2DaGuangZhou&",NFDT_3DaGuangZhou="&NFDT_3DaGuangZhou&",NFDT_4DaGuangZhou="&NFDT_4DaGuangZhou&",NFDT_1DaZhiBan='"&NFDT_1DaZhiBan&"',NFDT_2DaZhiBan='"&NFDT_2DaZhiBan&"',NFDT_3DaZhiBan='"&NFDT_3DaZhiBan&"',NFDT_4DaZhiBan='"&NFDT_4DaZhiBan&"',NFDT_FenSuoZhiBan='"&NFDT_FenSuoZhiBan&"',NFDT_GenXin='"&now()&"',NFDT_DongTai='"&NFDT_DongTai&"',NFDT_BaoGaoRen='"&NFDT_BaoGaoRen&"',NFDT_1DaDanDu="&NFDT_1DaDanDu&",NFDT_2DaDanDu="&NFDT_2DaDanDu&",NFDT_3DaDanDu="&NFDT_3DaDanDu&",NFDT_4DaDanDu="&NFDT_4DaDanDu&",NFDT_1DaMinJing="&NFDT_1DaMinJing&",NFDT_2DaMinJing="&NFDT_2DaMinJing&",NFDT_3DaMinJing="&NFDT_3DaMinJing&",NFDT_4DaMinJing="&NFDT_4DaMinJing&" where NFDT_FenSuo="&IID
'response.write sql
'response.end
Lours.Open sql,Louconn,1,3
call writerz(IID,NFDT_FenSuoZhiBan,NFDT_BaoGaoRen,NFDT_1DaZhiBan,NFDT_2DaZhiBan,NFDT_3DaZhiBan,NFDT_4DaZhiBan,NFDT_1DArenshu,NFDT_2DArenshu,NFDT_3DArenshu,NFDT_4DArenshu,NFDT_1DaWeiZhong,NFDT_2DaWeiZhong,NFDT_3DaWeiZhong,NFDT_4DaWeiZhong,NFDT_1DaGuangZhou,NFDT_2DaGuangZhou,NFDT_3DaGuangZhou,NFDT_4DaGuangZhou,NFDT_1DaDanDu,NFDT_2DaDanDu,NFDT_3DaDanDu,NFDT_4DaDanDu,NFDT_1DaMinJing,NFDT_2DaMinJing,NFDT_3DaMinJing,NFDT_4DaMinJing,NFDT_DongTai)
select case (IID)
case 1
response.write "<script language=JavaScript>" & chr(13) & "alert('动态信息更新成功');"&"window.location.href = '1fs.asp'"&" </script>" 
case 2
response.write "<script language=JavaScript>" & chr(13) & "alert('动态信息更新成功');"&"window.location.href = '2fs.asp'"&" </script>" 
case 3
response.write "<script language=JavaScript>" & chr(13) & "alert('动态信息更新成功');"&"window.location.href = '3fs.asp'"&" </script>" 
case 4
response.write "<script language=JavaScript>" & chr(13) & "alert('动态信息更新成功');"&"window.location.href = '4fs.asp'"&" </script>" 
case 5
response.write "<script language=JavaScript>" & chr(13) & "alert('动态信息更新成功');"&"window.location.href = '5fs.asp'"&" </script>" 
end select
set Lours=nothing
%>