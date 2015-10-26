
<%
'此文件为财付通支付配置文件

'☆★☆★☆★☆★☆★☆★财付通测试开关   0 关闭测试    1 开启测试☆★☆★☆★☆★☆★☆★

spid		= "12***701"							' 这里替换为您的实际商户号
sp_key		= "***"	' sp_key是32位商户密钥, 请替换为您的实际密钥
tenpay_dir	= "/tenpay_bargainor"					' WL 集成目录。

'☆★☆★☆★☆★☆★☆★财付通支付配置项。☆★☆★☆★☆★☆★☆★
domain				="http://www.myhomestay.com.cn"									'商户网站域名


site_name			="我的住家网安全支付"											'商户网站名称
imgtitle			="我们选择安全的财付通在线支付方式" 							'图片说明
imgsrc				=tenpay_dir&"/image/tenpay_buy.gif"								'图片地址

'mch_returl			= "http://www.myhomestay.com.cn/tenpay_myhomestay.com.cn/tenpay_notify.asp"			'通知URL
'show_url			= "http://www.myhomestay.com.cn/tenpay_myhomestay.com.cn/tenpay_show.asp"			'返回URL





' WL自己公司生成商户订单号函数 （ 原理：获取服务器 时间，格式MMDDHHMMSS ，10位长度）
Function CFTGetServerDate_BillNo 
	Dim strTmp, iYear,iMonth,iDate 
	'//iYear = Year(Date) 
	'//iMonth = Month(Date) 
	'//iDate = Day(Date) 
	
	iHour = Hour(Now)
	iMinute = Minute(Now)
	iSecond = Second(Now)
	
	'//strTmp = CStr(iYear)
	strTmp = Right(SESSION.SessionID, 4)		'WL
	
'//	If iMonth < 10 Then 
'//		strTmp = strTmp & "0" & Cstr(iMonth)
'//	Else 
'//		strTmp = strTmp & Cstr(iMonth)
'//	End If
'//	
'//	If iDate < 10 Then 
'//		strTmp = strTmp & "0" & Cstr(iDate) 
'//	Else 
'//		strTmp = strTmp & Cstr(iDate) 
'//	End If
'//	
	
	
	If iHour < 10 Then 
		strTmp = strTmp & "0" & Cstr(iHour) 
	Else 
		strTmp = strTmp & Cstr(iHour) 
	End If
	
	If iMinute < 10 Then 
		strTmp = strTmp & "0" & Cstr(iMinute) 
	Else 
		strTmp = strTmp & Cstr(iMinute) 
	End If
	
	If iSecond < 10 Then 
		strTmp = strTmp & "0" & Cstr(iSecond) 
	Else 
		strTmp = strTmp & Cstr(iSecond) 
	End If
	
	CFTGetServerDate_BillNo = strTmp 
End Function

%>