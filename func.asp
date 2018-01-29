<%
'危重、单独、所外留医非零改颜色'
function ChangeFontColor(text1)
	if(cint(text1))<>0 then
		response.write("<font color='red'>"&text1&"</font>")
	else
		response.write(text1)
	end if
end function
'检测时间是否为当天'
function Dantian(text1,time1)
''	response.write(time())
''	response.write(datediff("n","08:30:00",time()))
''	response.write(datediff("d",date(),time1))
	if(datediff("d",date(),time1))<>0 then
		if(datediff("n","08:30:00",time()))<0 then
			response.write(text1)
		else
			response.write("<font color='grey'>"&text1&"</font>")
		end if
	else
		response.write(text1)
''		response.write("已上报")
	end if
end function
'检测是否更新'
function NoUpdate(time1)
''	response.write(time1)
''	response.write(datediff("d",date(),time1))
	time2=datediff("n","08:30:00",time())
	if(datediff("d",date(),time1))<0 and datediff("n","08:30:00",time())>0 then
		response.write("交班时间已过"&time2&"分钟，还未上报！")
	else
''		response.write("已上报")
	end if
end function
'写入分所日志'
function writerz(aid,NFDT_SuoZhang,NFDT_ChuShi,NFDT_HuWeiDui,NFDT_YiYuan,NFDT_ShuiDianBan,NFDT_SiJi,NFDT_1DArenshu,NFDT_2DArenshu,NFDT_3DArenshu,NFDT_4DArenshu,NFDT_1DaWeiZhong,NFDT_2DaWeiZhong,NFDT_3DaWeiZhong,NFDT_4DaWeiZhong,NFDT_1DaGuangZhou,NFDT_2DaGuangZhou,NFDT_3DaGuangZhou,NFDT_4DaGuangZhou,NFDT_1DaDanDu,NFDT_2DaDanDu,NFDT_3DaDanDu,NFDT_4DaDanDu,NFDT_1DaMinJing,NFDT_2DaMinJing,NFDT_3DaMinJing,NFDT_4DaMinJing,NFDT_DongTai)
	sqlcmdd = ("select * from rizhi where datediff('d',rzdate,'"&date()&"') = 0")			'查询是否有当天记录
	set rss=server.createobject("adodb.recordset")
''	response.write(text)
	rss.Open sqlcmdd,Louconn,1,1
''		response.write(rss("rzdate"))
''		response.write(datediff("d","2018/1/10",date()))
''		response.write(sqlcmdd)
''		response.write(rss.recordcount)
''		response.end
		if rss.recordcount<>0 then

			a=rss("id")
			'判断有无当天记录'
			select case aid
				case 1
					text = ("update rizhi set 1fszb='"&NFDT_SuoZhang&"',1fsbgr='"&NFDT_ChuShi&"',1fs1dzb='"&NFDT_HuWeiDui&"',1fs2dzb='"&NFDT_YiYuan&"',1fs3dzb='"&NFDT_ShuiDianBan&"',1fs4dzb='"&NFDT_SiJi&"',1fs1drs="&NFDT_1DArenshu&",1fs2drs="&NFDT_2DArenshu&",1fs3drs="&NFDT_3DArenshu&",1fs4drs="&NFDT_4DArenshu&",1fs1dwz="&NFDT_1DaWeiZhong&",1fs2dwz="&NFDT_2DaWeiZhong&",1fs3dwz="&NFDT_3DaWeiZhong&",1fs4dwz="&NFDT_4DaWeiZhong&",1fs1dly="&NFDT_1DaGuangZhou&",1fs2dly="&NFDT_2DaGuangZhou&",1fs3dly="&NFDT_3DaGuangZhou&",1fs4dly="&NFDT_4DaGuangZhou&",1fs1ddd="&NFDT_1DaDanDu&",1fs2ddd="&NFDT_2DaDanDu&",1fs3ddd="&NFDT_3DaDanDu&",1fs4ddd="&NFDT_4DaDanDu&",1fs1dmj="&NFDT_1DaMinJing&",1fs2dmj="&NFDT_2DaMinJing&",1fs3dmj="&NFDT_3DaMinJing&",1fs4dmj="&NFDT_4DaMinJing&",1fsdt='"&NFDT_DongTai&"' where id="&a)	
'					response.write(text)
'					response.end		'
					set rs=server.createobject("adodb.recordset")
					rs.Open text,Louconn,1,3
					set rs=nothing
					rs.close
				case 2
					text = ("update rizhi set 2fszb='"&NFDT_SuoZhang&"',2fsbgr='"&NFDT_ChuShi&"',2fs1dzb='"&NFDT_HuWeiDui&"',2fs2dzb='"&NFDT_YiYuan&"',2fs3dzb='"&NFDT_ShuiDianBan&"',2fs4dzb='"&NFDT_SiJi&"',2fs1drs="&NFDT_1DArenshu&",2fs2drs="&NFDT_2DArenshu&",2fs3drs="&NFDT_3DArenshu&",2fs4drs="&NFDT_4DArenshu&",2fs1dwz="&NFDT_1DaWeiZhong&",2fs2dwz="&NFDT_2DaWeiZhong&",2fs3dwz="&NFDT_3DaWeiZhong&",2fs4dwz="&NFDT_4DaWeiZhong&",2fs1dly="&NFDT_1DaGuangZhou&",2fs2dly="&NFDT_2DaGuangZhou&",2fs3dly="&NFDT_3DaGuangZhou&",2fs4dly="&NFDT_4DaGuangZhou&",2fs1ddd="&NFDT_1DaDanDu&",2fs2ddd="&NFDT_2DaDanDu&",2fs3ddd="&NFDT_3DaDanDu&",2fs4ddd="&NFDT_4DaDanDu&",2fs1dmj="&NFDT_1DaMinJing&",2fs2dmj="&NFDT_2DaMinJing&",2fs3dmj="&NFDT_3DaMinJing&",2fs4dmj="&NFDT_4DaMinJing&",2fsdt='"&NFDT_DongTai&"' where id="&a)			
					set rs=server.createobject("adodb.recordset")
					rs.Open text,Louconn,1,3
					set rs=nothing
					rs.close
				case 3
					text = ("update rizhi set 3fszb='"&NFDT_SuoZhang&"',3fsbgr='"&NFDT_ChuShi&"',3fs1dzb='"&NFDT_HuWeiDui&"',3fs2dzb='"&NFDT_YiYuan&"',3fs3dzb='"&NFDT_ShuiDianBan&"',3fs4dzb='"&NFDT_SiJi&"',3fs1drs="&NFDT_1DArenshu&",3fs2drs="&NFDT_2DArenshu&",3fs3drs="&NFDT_3DArenshu&",3fs4drs="&NFDT_4DArenshu&",3fs1dwz="&NFDT_1DaWeiZhong&",3fs2dwz="&NFDT_2DaWeiZhong&",3fs3dwz="&NFDT_3DaWeiZhong&",3fs4dwz="&NFDT_4DaWeiZhong&",3fs1dly="&NFDT_1DaGuangZhou&",3fs2dly="&NFDT_2DaGuangZhou&",3fs3dly="&NFDT_3DaGuangZhou&",3fs4dly="&NFDT_4DaGuangZhou&",3fs1ddd="&NFDT_1DaDanDu&",3fs2ddd="&NFDT_2DaDanDu&",3fs3ddd="&NFDT_3DaDanDu&",3fs4ddd="&NFDT_4DaDanDu&",3fs1dmj="&NFDT_1DaMinJing&",3fs2dmj="&NFDT_2DaMinJing&",3fs3dmj="&NFDT_3DaMinJing&",3fs4dmj="&NFDT_4DaMinJing&",3fsdt='"&NFDT_DongTai&"' where id="&a)			
					set rs=server.createobject("adodb.recordset")
					rs.Open text,Louconn,1,3
					set rs=nothing
					rs.close
				case 4
					text = ("update rizhi set 4fszb='"&NFDT_SuoZhang&"',4fsbgr='"&NFDT_ChuShi&"',4fs1dzb='"&NFDT_HuWeiDui&"',4fs2dzb='"&NFDT_YiYuan&"',4fs3dzb='"&NFDT_ShuiDianBan&"',4fs4dzb='"&NFDT_SiJi&"',4fs1drs="&NFDT_1DArenshu&",4fs2drs="&NFDT_2DArenshu&",4fs3drs="&NFDT_3DArenshu&",4fs4drs="&NFDT_4DArenshu&",4fs1dwz="&NFDT_1DaWeiZhong&",4fs2dwz="&NFDT_2DaWeiZhong&",4fs3dwz="&NFDT_3DaWeiZhong&",4fs4dwz="&NFDT_4DaWeiZhong&",4fs1dly="&NFDT_1DaGuangZhou&",4fs2dly="&NFDT_2DaGuangZhou&",4fs3dly="&NFDT_3DaGuangZhou&",4fs4dly="&NFDT_4DaGuangZhou&",4fs1ddd="&NFDT_1DaDanDu&",4fs2ddd="&NFDT_2DaDanDu&",4fs3ddd="&NFDT_3DaDanDu&",4fs4ddd="&NFDT_4DaDanDu&",4fs1dmj="&NFDT_1DaMinJing&",4fs2dmj="&NFDT_2DaMinJing&",4fs3dmj="&NFDT_3DaMinJing&",4fs4dmj="&NFDT_4DaMinJing&",4fsdt='"&NFDT_DongTai&"' where id="&a)
					set rs=server.createobject("adodb.recordset")
					rs.Open text,Louconn,1,3
					set rs=nothing
					rs.close
			end select

		else
			kk=date()
			select case aid
				case 1
					text = ("insert into rizhi(1fszb,1fsbgr,1fs1dzb,1fs2dzb,1fs3dzb,1fs4dzb,1fs1drs,1fs2drs,1fs3drs,1fs4drs,1fs1dwz,1fs2dwz,1fs3dwz,1fs4dwz,1fs1dly,1fs2dly,1fs3dly,1fs4dly,1fs1ddd,1fs2ddd,1fs3ddd,1fs4ddd,1fs1dmj,1fs2dmj,1fs3dmj,1fs4dmj,1fsdt,rzdate)values('"&NFDT_SuoZhang&"','"&NFDT_ChuShi&"','"&NFDT_HuWeiDui&"','"&NFDT_YiYuan&"','"&NFDT_ShuiDianBan&"','"&NFDT_SiJi&"',"&NFDT_1DArenshu&","&NFDT_2DArenshu&","&NFDT_3DArenshu&","&NFDT_4DArenshu&","&NFDT_1DaWeiZhong&","&NFDT_2DaWeiZhong&","&NFDT_3DaWeiZhong&","&NFDT_4DaWeiZhong&","&NFDT_1DaGuangZhou&","&NFDT_2DaGuangZhou&","&NFDT_3DaGuangZhou&","&NFDT_4DaGuangZhou&","&NFDT_1DaDanDu&","&NFDT_2DaDanDu&","&NFDT_3DaDanDu&","&NFDT_4DaDanDu&","&NFDT_1DaMinJing&","&NFDT_2DaMinJing&","&NFDT_3DaMinJing&","&NFDT_4DaMinJing&",'"&NFDT_DongTai&"','"&kk&"')")			'
					Louconn.execute text
					rs.close
					set rs=nothing
				case 2
					text = ("insert into rizhi(2fszb,2fsbgr,2fs1dzb,2fs2dzb,2fs3dzb,2fs4dzb,2fs1drs,2fs2drs,2fs3drs,2fs4drs,2fs1dwz,2fs2dwz,2fs3dwz,2fs4dwz,2fs1dly,2fs2dly,2fs3dly,2fs4dly,2fs1ddd,2fs2ddd,2fs3ddd,2fs4ddd,2fs1dmj,2fs2dmj,2fs3dmj,2fs4dmj,2fsdt,rzdate)values('"&NFDT_SuoZhang&"','"&NFDT_ChuShi&"','"&NFDT_HuWeiDui&"','"&NFDT_YiYuan&"','"&NFDT_ShuiDianBan&"','"&NFDT_SiJi&"',"&NFDT_1DArenshu&","&NFDT_2DArenshu&","&NFDT_3DArenshu&","&NFDT_4DArenshu&","&NFDT_1DaWeiZhong&","&NFDT_2DaWeiZhong&","&NFDT_3DaWeiZhong&","&NFDT_4DaWeiZhong&","&NFDT_1DaGuangZhou&","&NFDT_2DaGuangZhou&","&NFDT_3DaGuangZhou&","&NFDT_4DaGuangZhou&","&NFDT_1DaDanDu&","&NFDT_2DaDanDu&","&NFDT_3DaDanDu&","&NFDT_4DaDanDu&","&NFDT_1DaMinJing&","&NFDT_2DaMinJing&","&NFDT_3DaMinJing&","&NFDT_4DaMinJing&",'"&NFDT_DongTai&"','"&kk&"')")			'
''					response.write(text)
''					response.end
					Louconn.execute text
					rs.close
					set rs=nothing
				case 3
					text = ("insert into rizhi(3fszb,3fsbgr,3fs1dzb,3fs2dzb,3fs3dzb,3fs4dzb,3fs1drs,3fs2drs,3fs3drs,3fs4drs,3fs1dwz,3fs2dwz,3fs3dwz,3fs4dwz,3fs1dly,3fs2dly,3fs3dly,3fs4dly,3fs1ddd,3fs2ddd,3fs3ddd,3fs4ddd,3fs1dmj,3fs2dmj,3fs3dmj,3fs4dmj,3fsdt,rzdate)values('"&NFDT_SuoZhang&"','"&NFDT_ChuShi&"','"&NFDT_HuWeiDui&"','"&NFDT_YiYuan&"','"&NFDT_ShuiDianBan&"','"&NFDT_SiJi&"',"&NFDT_1DArenshu&","&NFDT_2DArenshu&","&NFDT_3DArenshu&","&NFDT_4DArenshu&","&NFDT_1DaWeiZhong&","&NFDT_2DaWeiZhong&","&NFDT_3DaWeiZhong&","&NFDT_4DaWeiZhong&","&NFDT_1DaGuangZhou&","&NFDT_2DaGuangZhou&","&NFDT_3DaGuangZhou&","&NFDT_4DaGuangZhou&","&NFDT_1DaDanDu&","&NFDT_2DaDanDu&","&NFDT_3DaDanDu&","&NFDT_4DaDanDu&","&NFDT_1DaMinJing&","&NFDT_2DaMinJing&","&NFDT_3DaMinJing&","&NFDT_4DaMinJing&",'"&NFDT_DongTai&"','"&kk&"')")			'
					Louconn.execute text
					rs.close
					set rs=nothing
				case 4
					text = ("insert into rizhi(4fszb,4fsbgr,4fs1dzb,4fs2dzb,4fs3dzb,4fs4dzb,4fs1drs,4fs2drs,4fs3drs,4fs4drs,4fs1dwz,4fs2dwz,4fs3dwz,4fs4dwz,4fs1dly,4fs2dly,4fs3dly,4fs4dly,4fs1ddd,4fs2ddd,4fs3ddd,4fs4ddd,4fs1dmj,4fs2dmj,4fs3dmj,4fs4dmj,4fsdt,rzdate)values('"&NFDT_SuoZhang&"','"&NFDT_ChuShi&"','"&NFDT_HuWeiDui&"','"&NFDT_YiYuan&"','"&NFDT_ShuiDianBan&"','"&NFDT_SiJi&"',"&NFDT_1DArenshu&","&NFDT_2DArenshu&","&NFDT_3DArenshu&","&NFDT_4DArenshu&","&NFDT_1DaWeiZhong&","&NFDT_2DaWeiZhong&","&NFDT_3DaWeiZhong&","&NFDT_4DaWeiZhong&","&NFDT_1DaGuangZhou&","&NFDT_2DaGuangZhou&","&NFDT_3DaGuangZhou&","&NFDT_4DaGuangZhou&","&NFDT_1DaDanDu&","&NFDT_2DaDanDu&","&NFDT_3DaDanDu&","&NFDT_4DaDanDu&","&NFDT_1DaMinJing&","&NFDT_2DaMinJing&","&NFDT_3DaMinJing&","&NFDT_4DaMinJing&",'"&NFDT_DongTai&"','"&kk&"')")			'
					Louconn.execute text
					rs.close
					set rs=nothing
			end select
		end if
end function
'写入总所日志'
function writerzzs(NFDT_SuoZhang,NFDT_ChuShi,NFDT_HuWeiDui,NFDT_YiYuan,NFDT_ShuiDianBan,NFDT_SiJi,NFDT_DongTai)
	sqlcmdd = ("select * from rizhi where datediff('d',rzdate,'"&date()&"') = 0")			'查询有没当天记录
	set rss=server.createobject("adodb.recordset")
''	response.write(text)
	rss.Open sqlcmdd,Louconn,1,1
		if rss.recordcount<>0 then
		a=rss("id")
					text = ("update rizhi set szzb='"&NFDT_SuoZhang&"',cszb='"&NFDT_ChuShi&"',hwdzb='"&NFDT_HuWeiDui&"',yyzb='"&NFDT_YiYuan&"',sdbzb='"&NFDT_ShuiDianBan&"',sjzb='"&NFDT_SiJi&"',dzdt='"&NFDT_DongTai&"' where id="&a)			'更新记录
'		response.Write(text)
'		response.End()
					set rs=server.createobject("adodb.recordset")
					rs.Open text,Louconn,1,3
					set rs=nothing
					rs.close
		else
					tt=date()
					text = ("insert into rizhi(szzb,cszb,hwdzb,yyzb,sdbzb,sjzb,rzdate,dzdt)values('"&NFDT_SuoZhang&"','"&NFDT_ChuShi&"','"&NFDT_HuWeiDui&"','"&NFDT_YiYuan&"','"&NFDT_ShuiDianBan&"','"&NFDT_SiJi&"','"&tt&"','"&NFDT_DongTai&"')")		'插入新记录
''					response.write(date())
''					response.end
					Louconn.execute text
					rs.close
					set rs=nothing
		end if
end function
%>