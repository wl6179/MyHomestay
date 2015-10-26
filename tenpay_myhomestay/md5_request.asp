<!-- #include file="asp_md5.asp" -->
<!--#include file="tenpay_config.asp"-->

<html>
<title>MD5签名支付 Online</title>
<meta http-equiv="Cache-Control" content="no-cache"/> 
<body>
	<% Response.Charset="GB2312" %>
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
	cmdno		= "1"				' 财付通支付为"1" (当前只支持 cmdno=1)	
	bill_date	= CFTGetServerDate	' 交易日期 (yyyymmdd)	
	bank_type	= "0"				' 银行类型:	0		财付通
									'			1001	招商银行   
									'			1002	中国工商银行  
									'			1003	中国建设银行  
									'			1004	上海浦东发展银行   
									'			1005	中国农业银行  
									'			1006	中国民生银行  
									'			1008	深圳发展银行   
									'			1009	兴业银行   

	desc		= Request.Form("product_desc")		' 商品名称
	purchaser_id = ""				' 用户财付通帐号，如果没有可以置空
	bargainor_id = spid				' 商户号
	
	'--------------------Begin
	Dim strBillNo
	strBillNo = CFTGetServerDate_BillNo	'WL自己公司生成商户订单号函数
	'--------------------End
	
	'//sp_billno	 = strBillNo		' 商户生成的订单号(最多32位)	现在是10位
	sp_billno	 = Request.Form("str_BillNo")		' 商户生成的订单号(最多32位)	现在是10位	

	' 重要:
	' 交易单号(28位): 商户号(10位) + 日期(8位) + 流水号(10位), 必须按此格式生成, 且不能重复
	' 如果sp_billno超过10位, 则截取其中的流水号部分加到transaction_id后部(不足10位左补0)
	' 如果sp_billno不足10位, 则左补0, 加到transaction_id后部
	transaction_id = spid & bill_date & sp_billno 

	total_fee	 = Request.Form("total_money")					' 总金额, 分为单位
	fee_type	 = "1"					' 货币类型: 1 – RMB(人民币) 2 - USD(美元) 3 - HKD(港币)
	return_url	 = domain & tenpay_dir &"/notify_handler.asp"	 ' 财付通回调页面地址, (最长255个字符)
	attach		 = "WelcomeToHouseShop"	' 商户私有数据, 请求回调页面时原样返回

	' 生成MD5签名    
	sign_text = "cmdno=" & cmdno & "&date=" & bill_date & "&bargainor_id=" & bargainor_id &_
        "&transaction_id=" & transaction_id & "&sp_billno=" & sp_billno &_
        "&total_fee=" & total_fee & "&fee_type=" & fee_type & "&return_url=" & return_url &_
        "&attach=" & attach & "&key=" & sp_key

	md5_sign = UCase(ASP_MD5(sign_text))        ' 转换为大写 
	 
%>

<form name="form1" id="form1" action="https://www.tenpay.com/cgi-bin/v1.0/pay_gate.cgi">
	<input type=hidden name="cmdno"				value="<%=cmdno%>">
	<input type=hidden name="date"			    value="<%=bill_date%>">
	<input type=hidden name="bank_type"			value="<%=bank_type%>">
	<input type=hidden name="desc"				value="<%=desc%>">
	<input type=hidden name="purchaser_id"		value="<%=purchaser_id%>">
	<input type=hidden name="bargainor_id"		value="<%=bargainor_id%>">
	<input type=hidden name="transaction_id"	value="<%=transaction_id%>">
	<input type=hidden name="sp_billno"			value="<%=sp_billno%>">
	<input type=hidden name="total_fee"			value="<%=total_fee%>">
	<input type=hidden name="fee_type"			value="<%=fee_type%>">
	<input type=hidden name="return_url"		value="<%=return_url%>">
	<input type=hidden name="attach"			value="<%=attach%>">
	<input type=hidden name="sign"				value="<%=md5_sign%>">
	
	<!--<h4>发送 在线支付 Online</h4>
	<div>
	<textarea name="textarea1" cols="100" rows="6" readonly="true">
	cmdno=<%=cmdno%>&date=<%=bill_date%>&bank_type=<%=bank_type%>&
	desc=<%=desc%>&purchaser_id=<%=purchaser_id%>&
	bargainor_id=<%=bargainor_id%>&transaction_id=<%=transaction_id%>&
	sp_billno=<%=sp_billno%>&total_fee=<%=total_fee%>&
	fee_type=<%=fee_type%>&return_url=<%=return_url%>&
	attach=<%=attach%>&sign=<%=md5_sign%>
	</textarea>
	</div>-->
	<!--<input type=submit name=submit value=立刻支付Online />-->&nbsp;&nbsp;&nbsp;<!--<input type=button name=close value=关闭 onClick="window.close();" />-->
	
	
</form>
	<script language=javascript type="text/javascript">
		document.form1.submit();
	</script>
</body>
</html>
