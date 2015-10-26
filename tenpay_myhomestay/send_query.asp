<!-- #include file="asp_md5.asp" -->
<!--#include file="tenpay_config.asp"-->

<html>
<title>订单查询接口测试</title>
<meta http-equiv="Cache-Control" content="no-cache"/> 
<body>
<%
	' 获取服务器日期，格式YYYYMMDD
	Function CFTGetServerDate 
		Dim strTmp, iYear,iMonth,iDate 
		iYear = Year(Date) 
		iMonth = Month(Date) 
		iDate = Day(Date) 

		strTmp = CStr(iYear)
		If iMonth < 10 Then 
			strTmp = strTmp & "0" & Cstr(iMonth)
		Else 
			strTmp = strTmp & Cstr(iMonth)
		End If 
		If iDate < 10 Then 
			strTmp = strTmp & "0" & Cstr(iDate) 
		Else 
			strTmp = strTmp & Cstr(iDate) 
		End If 
		CFTGetServerDate = strTmp 
	End Function

	Dim request_text
	Dim md5_sign

	'< !--#include file="tenpay_config.asp"-- >
	
	
	' 下面是请求参数
	cmdno			= "2"					' 订单查询业务为2
	bill_date	= "20080516"	' 需要查询的交易日期 (yyyymmdd)		

	bargainor_id = spid					' 商户号
	sp_billno	 = "3771163914"		' 商户生成的订单号(最多32位) 要查询的商户订单号！	

	' 重要:
	' 需要查询的交易单号(28位): 商户号(10位) + 日期(8位) + 流水号(10位), 必须按此格式生成, 且不能重复
	' 如果sp_billno超过10位, 则截取其中的流水号部分加到transaction_id后部(不足10位左补0)
	' 如果sp_billno不足10位, 则左补0, 加到transaction_id后部
	transaction_id = spid & bill_date & sp_billno 
	
	return_url = domain & tenpay_dir &"/query_result.asp" ' 财付通回调页面地址,通知订单信息,推荐使用ip地址的方式(最长255个字符)
	attach		 = "Welcome6"	' 商户私有数据, 请求回调页面时原样返回

	' 生成MD5签名    
	sign_text = "cmdno=" & cmdno & "&date=" & bill_date & "&bargainor_id=" & bargainor_id &_
        "&transaction_id=" & transaction_id & "&sp_billno=" & sp_billno &_
        "&return_url=" & return_url &_
        "&attach=" & attach & "&key=" & sp_key

  md5_sign = UCase(ASP_MD5(sign_text))        ' 转换为大写
  
%>

<form action="http://portal.tenpay.com/cfbiportal/cgi-bin/cfbiqueryorder.cgi">
<input type=hidden name="cmdno"				value="<%=cmdno%>">
<input type=hidden name="date"			    value="<%=bill_date%>">
<input type=hidden name="bargainor_id"		value="<%=bargainor_id%>">
<input type=hidden name="transaction_id"	value="<%=transaction_id%>">
<input type=hidden name="sp_billno"			value="<%=sp_billno%>">
<input type=hidden name="return_url"		value="<%=return_url%>">
<input type=hidden name="attach"			value="<%=attach%>">
<input type=hidden name="sign"				value="<%=md5_sign%>">

<h4>发送给财付通的请求串</h4>
<div>
<textarea name="textarea1" cols="100" rows="6" readonly="true">
cmdno=<%=cmdno%>&date=<%=bill_date%>&bargainor_id=<%=bargainor_id%>&
transaction_id=<%=transaction_id%>&sp_billno=<%=sp_billno%>&
return_url=<%=return_url%>&attach=<%=attach%>&sign=<%=md5_sign%>
</textarea>
</div>
<input type=submit name=submit value=查询订单>
</form>
</body>

</html>
