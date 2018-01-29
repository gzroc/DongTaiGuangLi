function sub(){
	var aa=document.getElementById("NFDT_BaoGaoRen").value
	if(document.getElementById("NFDT_1DArenshu").value.length==0 || document.getElementById("NFDT_1DaWeiZhong").value.length==0 || document.getElementById("NFDT_1DaGuangZhou").value.length==0 || document.getElementById("NFDT_1DaZhiBan").value.length==0 || document.getElementById("NFDT_2DArenshu").value.length==0 || document.getElementById("NFDT_2DaWeiZhong").value.length==0 || document.getElementById("NFDT_2DaGuangZhou").value.length==0 || document.getElementById("NFDT_2DaZhiBan").value.length==0 || document.getElementById("NFDT_3DArenshu").value.length==0 || document.getElementById("NFDT_3DaWeiZhong").value.length==0 || document.getElementById("NFDT_3DaGuangZhou").value.length==0 || document.getElementById("NFDT_3DaZhiBan").value.length==0 || document.getElementById("NFDT_4DArenshu").value.length==0 || document.getElementById("NFDT_4DaWeiZhong").value.length==0 || document.getElementById("NFDT_4DaGuangZhou").value.length==0 || document.getElementById("NFDT_4DaZhiBan").value.length==0 || document.getElementById("NFDT_FenSuoZhiBan").value.length==0)
	{
		alert("请填写全部信息。");
		return false;
	}
	if(aa.length==0)
	{
		alert("报告人姓名必须填写，谢谢！");
		return false;
	}
	if(confirm("确认更新信息吗？"))
	{
		form1.submit();
	}
	else
	{
		return false;
	}
} 
function sub2(){
	var aa=document.getElementById("NFDT_BaoGaoRen").value
	if(document.getElementById("NFDT_1DArenshu").value.length==0 || document.getElementById("NFDT_1DaWeiZhong").value.length==0 || document.getElementById("NFDT_1DaGuangZhou").value.length==0 || document.getElementById("NFDT_1DaZhiBan").value.length==0 || document.getElementById("NFDT_2DArenshu").value.length==0 || document.getElementById("NFDT_2DaWeiZhong").value.length==0 || document.getElementById("NFDT_2DaGuangZhou").value.length==0 || document.getElementById("NFDT_2DaZhiBan").value.length==0 || document.getElementById("NFDT_3DArenshu").value.length==0 || document.getElementById("NFDT_3DaWeiZhong").value.length==0 || document.getElementById("NFDT_3DaGuangZhou").value.length==0 || document.getElementById("NFDT_3DaZhiBan").value.length==0 || document.getElementById("NFDT_4DArenshu").value.length==0 || document.getElementById("NFDT_4DaWeiZhong").value.length==0 || document.getElementById("NFDT_4DaGuangZhou").value.length==0 || document.getElementById("NFDT_4DaZhiBan").value.length==0 || document.getElementById("NFDT_FenSuoZhiBan").value.length==0)
	{
		alert("请填写全部信息。");
		return false;
	}
	if(aa.length==0)
	{
		alert("报告人姓名必须填写，谢谢！");
		return false;
	}
	if(confirm("确认更新信息吗？"))
	{
		form2.submit();
	}
	else
	{
		return false;
	}
} 

function sub3()
{
//	alert("hello");
	if(document.getElementById("NFDT_SuoZhang").value.length==0 || document.getElementById("NFDT_ChuShi").value.length==0 || document.getElementById("NFDT_HuWeiDui").value.length==0 || document.getElementById("NFDT_YiYuan").value.length==0 )
	{
		alert("请填写全部信息。");
		return false;
	}
	if(confirm("确认上报信息吗？"))
	{
		form1.submit();
	}
	else
	{
		return false;
	}
} 
